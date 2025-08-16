import { argbToHex, getElementByName, getElementsByName, hexToRgb, rgbToHex } from "../utils";
import { Drawing, DrawingObjectType, DrawingPosition, MediaFile } from "./types";
import { Theme } from "../theme/types";

export const parseDrawingXml = (drawingStr: string, relStr: string, media: MediaFile[], themes: Theme[] = []): Drawing[] => {
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
            ...getElementsByName(drawingElement, 'xdr:absoluteAnchor'),
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
            const anchorType = anchor.tagName.includes('oneCellAnchor')
                ? 'oneCell'
                : (anchor.tagName.includes('absoluteAnchor') ? 'absolute' : 'twoCell');
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

            // Fallback: explicit textbox element (xdr:txbx)
            if (!containerElement) {
                const txbx = getElementByName(anchor, 'xdr:txbx');
                if (txbx) {
                    drawingType = 'textbox';
                    containerElement = txbx;
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
            const ext = getElementByName(anchor, 'xdr:ext'); // size for oneCell/absolute
            const pos = getElementByName(anchor, 'xdr:pos'); // absolute pos

            const getPosition = (posElement: Element | undefined) => ({
                col: parseInt(getElementByName(posElement, 'xdr:col')?.textContent || '0', 10),
                colOff: parseInt(getElementByName(posElement, 'xdr:colOff')?.textContent || '0', 10),
                row: parseInt(getElementByName(posElement, 'xdr:row')?.textContent || '0', 10),
                rowOff: parseInt(getElementByName(posElement, 'xdr:rowOff')?.textContent || '0', 10),
            });

            const position: DrawingPosition = {
                from: getPosition(from),
                to: to ? getPosition(to) : getPosition(from),
            };

            let sizeEMU = (anchorType === 'oneCell' || anchorType === 'absolute') && ext
                ? {
                    cx: parseInt(ext.getAttribute('cx') || '0', 10),
                    cy: parseInt(ext.getAttribute('cy') || '0', 10),
                }
                : undefined;
            const absEMU = anchorType === 'absolute' && pos
                ? {
                    x: parseInt(pos.getAttribute('x') || '0', 10),
                    y: parseInt(pos.getAttribute('y') || '0', 10),
                }
                : undefined;

            // Try to override size from shape transform if present
            if (containerElement) {
                const xfrm = getElementByName(containerElement, 'a:xfrm');
                const xfrmExt = getElementByName(xfrm, 'a:ext');
                if (xfrmExt) {
                    const cx = parseInt(xfrmExt.getAttribute('cx') || '0', 10);
                    const cy = parseInt(xfrmExt.getAttribute('cy') || '0', 10);
                    if (cx > 0 || cy > 0) {
                        // prefer explicit transform ext when available
                        sizeEMU = { cx: cx || 0, cy: cy || 0 };
                    }
                }
            }

            const properties = extractPropertiesByType(drawingType, containerElement, themes);

            drawings.push({
                id,
                name,
                title,
                description,
                type: drawingType,
                base64: mediaFile?.base64 || '',
                anchorType,
                sizeEMU,
                absEMU,
                properties,
                position,
            });
        });
    }

    return drawings;
};

function extractPropertiesByType(type: DrawingObjectType, container: Element | undefined, themes: Theme[]): Record<string, any> {
    if (!container) return {};

    const colorFrom = (root?: Element): string => {
        if (!root) return '';
        const srgb = getElementByName(root, 'a:srgbClr');
        if (srgb) return `#${srgb.getAttribute('val')}`;
        const scheme = getElementByName(root, 'a:schemeClr');
        if (scheme) {
            const name = scheme.getAttribute('val') || '';
            const t = themes.find(th => th.name === name);
            let hex = t?.val ? argbToHex(t.val) : '';
            const shade = getElementByName(scheme, 'a:shade')?.getAttribute('val');
            const tint = getElementByName(scheme, 'a:tint')?.getAttribute('val');
            if (hex) {
                const rgb = hexToRgb(hex);
                if (shade) {
                    const f = 1 - (+shade / 100000);
                    rgb.r = Math.round(rgb.r * f);
                    rgb.g = Math.round(rgb.g * f);
                    rgb.b = Math.round(rgb.b * f);
                }
                if (tint) {
                    const f = +tint / 100000;
                    rgb.r = Math.round(rgb.r * (1 - f) + 255 * f);
                    rgb.g = Math.round(rgb.g * (1 - f) + 255 * f);
                    rgb.b = Math.round(rgb.b * (1 - f) + 255 * f);
                }
                hex = rgbToHex(rgb.r, rgb.g, rgb.b);
            }
            return hex;
        }
        const scrgb = getElementByName(root, 'a:scrgbClr');
        if (scrgb) {
            const r = parseFloat(scrgb.getAttribute('r') || '0');
            const g = parseFloat(scrgb.getAttribute('g') || '0');
            const b = parseFloat(scrgb.getAttribute('b') || '0');
            const to255 = (v: number) => (v <= 1 ? Math.round(v * 255) : Math.round(v));
            return rgbToHex(to255(r), to255(g), to255(b));
        }
        return '';
    };

    const shapeStyles = (sp: Element) => {
        const spPr = getElementByName(sp, 'xdr:spPr');
        const solidFill = getElementByName(spPr, 'a:solidFill');
        let fillColor = colorFrom(solidFill);
        const ln = getElementByName(spPr, 'a:ln');
        const lnFill = getElementByName(ln, 'a:solidFill');
        let lineColor = colorFrom(lnFill);
        let lineWidth = ln?.getAttribute('w') || '';
        const style = getElementByName(sp, 'xdr:style');
        if (!fillColor) {
            const fillRef = getElementByName(style, 'a:fillRef');
            fillColor = colorFrom(fillRef);
        }
        if (!lineColor) {
            const lnRef = getElementByName(style, 'a:lnRef');
            lineColor = colorFrom(lnRef);
        }
        return { fillColor, lineColor, lineWidth };
    };

    switch (type) {
        case 'image':
            {
            const blip = getElementByName(container, 'a:blip');
            return {
                embedId: blip?.getAttribute('r:embed') || ''
            };
        }
        case 'textbox':
            {
            const textElements = getElementsByName(container, 'a:t');
            const text = textElements.map(el => el.textContent).join('\n');
            const rPr = getElementByName(container, 'a:rPr');
            const sz = rPr?.getAttribute('sz');
            const fontPt = sz ? (+sz / 100) : undefined;
            const bold = rPr?.getAttribute('b') === '1' || rPr?.getAttribute('b') === 'true';
            const italic = rPr?.getAttribute('i') === '1' || rPr?.getAttribute('i') === 'true';
            let txtColor = colorFrom(getElementByName(rPr, 'a:solidFill'));
            // Fallback to style fontRef when run color is not set
            if (!txtColor) {
                const style = getElementByName(container, 'xdr:style');
                const fontRef = getElementByName(style, 'a:fontRef');
                txtColor = colorFrom(fontRef);
            }
            const pPr = getElementByName(container, 'a:pPr');
            const algn = pPr?.getAttribute('algn') || '';
            const textAlign = algn === 'ctr' ? 'center' : (algn === 'r' ? 'right' : 'left');
            const { fillColor, lineColor, lineWidth } = shapeStyles(container);
            return { text, fontPt, bold, italic, textColor: txtColor, textAlign, fillColor, lineColor, lineWidth };
            }
        case 'shape':
            {
            const prstGeom = getElementByName(container, 'a:prstGeom');
            const shapeType = prstGeom?.getAttribute('prst') || 'custom';
            const { fillColor, lineColor, lineWidth } = shapeStyles(container);
            return { fillColor, shapeType, lineColor, lineWidth };
            }
        case 'connector':
            {
            const line = getElementByName(container, 'a:ln');
            const lineColor = colorFrom(getElementByName(line, 'a:solidFill'));
            const lineWidth = line?.getAttribute('w') || '';
            return { lineColor, lineWidth };
            }
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
