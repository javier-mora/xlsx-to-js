import { getElementByName, getElementsByName } from "../utils";
import { Workbook } from "./types";

export const parseWorkbookXml = (str: string): Workbook => {
    const workbook: Workbook = {
        workSheets: [],
    }; 
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const workbookElement = getElementByName(xmlDoc, 'workbook');
    const sheetsElement = getElementByName(workbookElement, 'sheets');
    const sheetsArray = getElementsByName(sheetsElement, 'sheet');

    if (sheetsArray) {
        sheetsArray.forEach(x => {
            workbook.workSheets.push({
                id: +(x.getAttribute('sheetId') ?? 0),
                name: x.getAttribute('name') ?? '',
                dimention: '',
                data: [],
                columnStyles: [],
                rowStyles: [],
                mergeCells: [],
            });
        });
    }

    return workbook;
};