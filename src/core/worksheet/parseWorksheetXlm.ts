import { StyleSheet } from "../style/types";
import { getElementByName, getElementsByName, getPositionInArray, getRangeArray } from "../utils";
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
    };
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const worksheetElement = getElementByName(xmlDoc, 'worksheet');
    const dimensionElement = getElementByName(worksheetElement, 'dimension');
    const sheetDataElement = getElementByName(worksheetElement, 'sheetData');
    const mergeCellsElement = getElementByName(worksheetElement, 'mergeCells');
    const colsElement = getElementByName(worksheetElement, 'cols');
    const colsArray = getElementsByName(colsElement, 'col');
    const rowsArray = getElementsByName(sheetDataElement, 'row');
    const mergeCellArray = getElementsByName(mergeCellsElement, 'mergeCell');

    if (dimensionElement) {
        worksheet.dimention = dimensionElement.getAttribute('ref') ?? '';
        worksheet.data = getRangeArray(
            worksheet.dimention,
            dense ? { ref: '', value: '', formula: '' } : undefined
        );
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
                        value: t !== 's' ? value : sharedStrings[+value],
                        formula: formula,
                        style: (s !== -1  && styleSheet.cells.length > 0)
                            ? {
                                bgColor: styleSheet.fills[styleSheet.cells[s].fillId].bgColor,
                                fgColor: styleSheet.fills[styleSheet.cells[s].fillId].fgColor,
                                fontName: styleSheet.fonts[styleSheet.cells[s].fontId].name,
                                fontSize: styleSheet.fonts[styleSheet.cells[s].fontId].size,
                                vAlign: '',
                                hAlign: '',
                                border: styleSheet.borders[styleSheet.cells[s].borderId],
                            }
                            : undefined,
                    }
                });
            }
        });

        if (mergeCellArray) {
            mergeCellArray.forEach(x => {
                worksheet.mergeCells.push(x.getAttribute('ref') ?? '');
            });
        }
    }
    
    return worksheet;
}