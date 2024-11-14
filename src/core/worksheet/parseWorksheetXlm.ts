import { StyleSheet } from "../style/types";
import { excelSerialToJSDate, getElementByName, getElementsByName, getPositionInArray, getRangeArray, getSheetDimension } from "../utils";
import { ColStyle, WorkSheet } from "./types";

export const parseWorksheetXml = (str: string, styleSheet: StyleSheet, sharedStrings: string[], dense: boolean): WorkSheet => {
    const worksheet: WorkSheet = {
        id: 0,
        name: '',
        dimention: '',
        data: [],
        columnStyles: [],
        rowStyles: [],
        mergeCells: [],
        defaultColWidth: 12.5,
        defaultRowHeight: 3.25,
    };
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const worksheetElement = getElementByName(xmlDoc, 'worksheet');
    const dimensionElement = getElementByName(worksheetElement, 'dimension');
    const sheetDataElement = getElementByName(worksheetElement, 'sheetData');
    const mergeCellsElement = getElementByName(worksheetElement, 'mergeCells');
    const colsElement = getElementByName(worksheetElement, 'cols');
    const sheetFormatPrElement = getElementByName(worksheetElement, 'sheetFormatPr');
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
            if (index >= 0) {
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
            }
            // Row style
            worksheet.rowStyles.push({
                height: +(x.getAttribute('ht') ?? worksheet.defaultRowHeight),
            });
        });

        if (mergeCellArray) {
            mergeCellArray.forEach(x => {
                worksheet.mergeCells.push(x.getAttribute('ref') ?? '');
            });
        }
    }
    
    return worksheet;
}