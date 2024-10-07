import { getElementByName, getElementsByName } from "../utils";


export const parseSharedStringsXml = (str: string): string[] => {
    const strings: string[] = [];
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const sstElement = getElementByName(xmlDoc, 'sst');
    const tArray = getElementsByName(sstElement, 't');

    tArray.forEach(x => {
        strings.push(x.textContent ?? '');
    });

    return strings;
};