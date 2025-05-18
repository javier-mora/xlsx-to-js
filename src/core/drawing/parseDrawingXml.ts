import { argbToHex, getElementByName, getElementsByName, positionFromExt } from "../utils";
import { Drawing, DrawingObjectType, DrawingPosition, MediaFile } from "./types";

export const parseDrawingXml = (drawingStr: string, relStr: string, media: MediaFile[]): Drawing[] => {
    const drawings: Drawing[] = [];
    const rels: {id: string, uri: string}[] = [];

    const xmlDoc = new DOMParser().parseFromString(drawingStr, 'text/xml');
    const xmlRel = new DOMParser().parseFromString(relStr, 'text/xml');
    
    const relElement = getElementByName(xmlRel, 'Relationships');
    const drawingElement = getElementByName(xmlDoc, 'xdr:wsDr');
    
    const relationshipArray = getElementsByName(relElement, 'Relationship');
    const drawingsArray = drawingElement
        ? [
            ...getElementsByName(drawingElement, 'xdr:twoCellAnchor'),
            ...getElementsByName(drawingElement, 'xdr:oneCellAnchor'),
          ]
        : [];

    if (relationshipArray) {
        relationshipArray.forEach(rel => {
            rels.push({
                id: rel.getAttribute('Id') || '',
                uri: rel.getAttribute('Target') || '',
            });
        });
    }

    if (drawingsArray) {
        drawingsArray.forEach(anchor => {
            const elementTypeMap: { [tag: string]: DrawingObjectType } = {
                'xdr:pic': 'image',
                'xdr:sp': 'shape',
                'xdr:cxnSp': 'connector',
                'xdr:grpSp': 'group'
            };

            let drawingType: DrawingObjectType = 'unknown';
            let containerElement: Element | undefined;

            for (const [tag, detectedType] of Object.entries(elementTypeMap)) {
                const foundElement = getElementByName(anchor, tag);
                if (foundElement) {
                    drawingType = detectedType;
                    containerElement = foundElement;

                    if (drawingType === 'shape' && getElementByName(containerElement, 'xdr:txBody')) {
                        drawingType = 'textbox';
                    }
                    break;
                }
            }

            // Obtener metadata del dibujo
            const cNvPr = getElementByName(anchor, 'xdr:cNvPr');
            const id = cNvPr?.getAttribute('id') || '';
            const name = cNvPr?.getAttribute('name') || '';
            const title = cNvPr?.getAttribute('title') || '';
            const description = cNvPr?.getAttribute('descr') || '';

            // Obtener la imagen asociada
            const blip = getElementByName(anchor, 'a:blip');
            const embedId = blip?.getAttribute('r:embed') || '';
            const rel = rels.find(r => r.id === embedId);
            const mediaFileName = rel ? rel.uri.split('/').pop() : '';
            const mediaFile = media.find(m => m.name === mediaFileName);

            // Posiciones
            const from = getElementByName(anchor, 'xdr:from');
            const to = getElementByName(anchor, 'xdr:to');
            const ext = getElementByName(anchor, 'xdr:ext'); // Solo en oneCellAnchor

            const getPosition = (posElement: Element | undefined) => ({
                col: parseInt(getElementByName(posElement, 'xdr:col')?.textContent || '0', 10),
                colOff: parseInt(getElementByName(posElement, 'xdr:colOff')?.textContent || '0', 10),
                row: parseInt(getElementByName(posElement, 'xdr:row')?.textContent || '0', 10),
                rowOff: parseInt(getElementByName(posElement, 'xdr:rowOff')?.textContent || '0', 10),
            });

            const position: DrawingPosition = {
                from: getPosition(from),
                to: to
                    ? getPosition(to)
                    : {
                        col: positionFromExt(getPosition(from), ext?.getAttribute('cx')),
                        colOff: 0,
                        row: positionFromExt(getPosition(from), ext?.getAttribute('cy')),
                        rowOff: 0,
                    },
            };

            const properties = extractPropertiesByType(drawingType, containerElement);

            drawings.push({
                id,
                name,
                title,
                description,
                type: drawingType,
                base64: mediaFile?.base64 || '',
                properties,
                position,
            });
        });
    }

    return drawings;
};

function extractPropertiesByType(type: DrawingObjectType, container?: Element): Record<string, any> {
    if (!container) return {};

    switch (type) {
        case 'image':
            const blip = getElementByName(container, 'a:blip');
            return {
                embedId: blip?.getAttribute('r:embed') || ''
            };

        case 'textbox':
            const textElements = getElementsByName(container, 'a:t');
            const text = textElements.map(el => el.textContent).join('\n');
            return { text };

        case 'shape':
            const solidFill = getElementByName(container, 'a:solidFill');
            const fillColor = argbToHex(extractColor(solidFill) ?? '');
            const prstGeom = getElementByName(container, 'a:prstGeom');
            const shapeType = prstGeom?.getAttribute('prst') || 'custom';
            return { fillColor, shapeType };

        case 'connector':
            const line = getElementByName(container, 'a:ln');
            const lineColor = argbToHex(extractColor(getElementByName(line, 'a:solidFill')) ?? '');
            const lineWidth = line?.getAttribute('w') || '';
            return { lineColor, lineWidth };

        case 'group':
            return {};

        default:
            return {};
    }
}

function extractColor(colorElement?: Element): string | undefined {
    if (!colorElement) return undefined;

    const srgbClr = getElementByName(colorElement, 'a:srgbClr');
    if (srgbClr) return `#${srgbClr.getAttribute('val')}`;

    const schemeClr = getElementByName(colorElement, 'a:schemeClr');
    if (schemeClr) return schemeClr.getAttribute('val') ?? undefined;

    return undefined;
}