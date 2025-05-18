import { DrawingFile } from "../drawing/types";
import { StyleSheet } from "../style/types";
import { excelSerialToJSDate, getElementByName, getElementsByName, getPositionInArray, getRangeArray, getSheetDimension } from "../utils";
import { ColStyle, WorkSheet } from "./types";

export const parseWorksheetXml = (str: string, relStr: string | undefined, styleSheet: StyleSheet, sharedStrings: string[], drawingFiles: DrawingFile[], dense: boolean, skipHiddenRows: boolean): WorkSheet => {
    const worksheet: WorkSheet = {
        id: 0,
        name: '',
        dimention: '',
        data: [],
        columnStyles: [],
        rowStyles: [],
        mergeCells: [],
        drawings: [],
        defaultColWidth: 12.5,
        defaultRowHeight: 3.25,
        zeroHeight: false,
    };
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const xmlRel = relStr ? new DOMParser().parseFromString(relStr, 'text/xml') : undefined;
    const worksheetElement = getElementByName(xmlDoc, 'worksheet');
    const dimensionElement = getElementByName(worksheetElement, 'dimension');
    const sheetDataElement = getElementByName(worksheetElement, 'sheetData');
    const mergeCellsElement = getElementByName(worksheetElement, 'mergeCells');
    const colsElement = getElementByName(worksheetElement, 'cols');
    const sheetFormatPrElement = getElementByName(worksheetElement, 'sheetFormatPr');
    const drawingElement = getElementsByName(xmlRel, 'Relationship');
    const colsArray = getElementsByName(colsElement, 'col');
    const rowsArray = getElementsByName(sheetDataElement, 'row');
    const mergeCellArray = getElementsByName(mergeCellsElement, 'mergeCell');
    

    if (sheetDataElement && !dimensionElement) {
        worksheet.dimention = getSheetDimension(sheetDataElement);
    }

    if (dimensionElement) {
        worksheet.dimention = dimensionElement.getAttribute('ref') ?? '';
    }

    if (worksheet.dimention !== '') {
        worksheet.data = getRangeArray(
            worksheet.dimention,
            dense ? { ref: '', value: '', formula: '' } : undefined
        );
    }

    if (sheetFormatPrElement) {
        const baseColWidth = +(sheetFormatPrElement.getAttribute('baseColWidth') ?? 10);
        worksheet.defaultColWidth = +(sheetFormatPrElement.getAttribute('defaultColWidth') ?? (baseColWidth + 5));
        worksheet.defaultRowHeight = +(sheetFormatPrElement.getAttribute('defaultRowHeight') ?? 3.25);
        worksheet.zeroHeight = sheetFormatPrElement.getAttribute('zeroHeight') === "1";
    }

    if (colsArray) {
        const columnStyles: ColStyle[] = [];
        colsArray.forEach(x => {
            columnStyles.push({
                min: +(x.getAttribute('min') ?? 1),
                max: +(x.getAttribute('max') ?? 1),
                width: +(x.getAttribute('width') ?? 1),
                hidden: (x.getAttribute('hidden') ?? 'false') === 'true',
                collapsed: (x.getAttribute('collapsed') ?? 'false') === 'true',
            });
        });
        worksheet.columnStyles = columnStyles;
    }

    if (rowsArray) {
        rowsArray.forEach(x => {
            const index = +(x.getAttribute('r') ?? 0) - 1;
            const hiddenProp = (x.getAttribute('hidden') ?? 'false') === 'true';
            const collapsedProp = (x.getAttribute('collapsed') ?? 'false') === 'true';

            if (index >= 0 && (!skipHiddenRows || (!hiddenProp && !collapsedProp))) {
                const cols = getElementsByName(x, 'c');
                cols.forEach(y => {
                    const r = y.getAttribute('r') ?? 'A'; // Row
                    const s = +(y.getAttribute('s') ?? -1); // Style
                    const t = y.getAttribute('t') ?? ''; // Type
                    const formula = getElementByName(y, 'f')?.textContent ?? '';
                    const value = getElementByName(y, 'v')?.textContent ?? '';

                    const pos = getPositionInArray(r);
                    worksheet.data[pos.row][pos.col] = {
                        ref: r,
                        value: t !== 's' 
                            ? (
                                (value !== '' && !isNaN(+value) && s !== -1  && styleSheet.cells.length > 0 && styleSheet.cells[s].numFmtId === 14)
                                    ? excelSerialToJSDate(+value).toLocaleDateString() // Date
                                    : value // Number
                            ) 
                            : sharedStrings[+value],
                        formula: formula,
                        style: (s !== -1  && styleSheet.cells.length > 0)
                            ? {
                                bgColor: styleSheet.fills[styleSheet.cells[s].fillId].bgColor,
                                fgColor: styleSheet.fills[styleSheet.cells[s].fillId].fgColor,
                                fontName: styleSheet.fonts[styleSheet.cells[s].fontId].name,
                                fontSize: styleSheet.fonts[styleSheet.cells[s].fontId].size,
                                fontColor: styleSheet.fonts[styleSheet.cells[s].fontId].color,
                                bold: styleSheet.fonts[styleSheet.cells[s].fontId].bold,
                                italic: styleSheet.fonts[styleSheet.cells[s].fontId].italic,
                                vAlign: styleSheet.cells[s].alignment?.vertical ?? 'bottom',
                                hAlign: styleSheet.cells[s].alignment?.horizontal ?? '',
                                wrapText: styleSheet.cells[s].alignment?.wrapText ?? false,
                                border: styleSheet.borders[styleSheet.cells[s].borderId],
                            }
                            : undefined,
                    }

                    if (worksheet.data[pos.row][pos.col]?.style?.hAlign === '') {
                        worksheet.data[pos.row][pos.col]!.style!.hAlign = isNaN(+worksheet.data[pos.row][pos.col].value) ? 'left' : 'right';
                    }
                });

                // Row style
                if (!skipHiddenRows || (cols.length > 0)) {
                    worksheet.rowStyles.push({
                        r: +(x.getAttribute('r') ?? 0),
                        height: +(x.getAttribute('ht') ?? worksheet.defaultRowHeight),
                        hidden: (x.getAttribute('hidden') ?? 'false') === 'true',
                        collapsed: (x.getAttribute('collapsed') ?? 'false') === 'true',
                    });
                }
            }
        });

        if (mergeCellArray) {
            mergeCellArray.forEach(x => {
                worksheet.mergeCells.push(x.getAttribute('ref') ?? '');
            });
        }
    }

    if (skipHiddenRows) {
        if (worksheet.defaultRowHeight === 0 && worksheet.zeroHeight) {
            worksheet.data = worksheet.data.filter(x => x.length > 0 && x[0] !== undefined && x.some(y => y.ref !== ''));
        }
    }

    const drawingRel = drawingElement.find(rel =>
        rel.getAttribute('Type')?.includes('drawing')
    );

    if (drawingRel) {
        const drawingTarget = drawingRel.getAttribute('Target') || '';
        const drawingFileName = drawingTarget.split('/').pop() || '';

        const drawingFile = drawingFiles.find(df => df.src === drawingFileName);
        worksheet.drawings = drawingFile?.drawings ?? [];
    }
    
    return worksheet;
}