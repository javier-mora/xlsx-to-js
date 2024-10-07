import JSZip from "jszip";
import { XlsxParserOptions } from "./types";
import { parseThemeXml, Theme } from "./core/theme";
import { parseStylesXml, StyleSheet } from "./core/style";
import { parseWorkbookXml, Workbook } from "./core/workbook";
import { parseSharedStringsXml } from "./core/sharedString";
import { parseWorksheetXml } from "./core/worksheet";

export class XlsxParser {
    
    /**
     * Extract data from spreadsheet bytes stored in a `ArrayBuffer`.
     * @param file Workbook file
     * @param options Parser options
     * @returns 
     */
    async readFile (file: ArrayBuffer, options: XlsxParserOptions = { dense: false, styles: false }) {
        const nZip = new JSZip();
        const result = await nZip.loadAsync(file);
        const files: string[] = [];
        result.forEach((p) => files.push(p));
        const sheetFiles = files.filter(x => x.includes('xl/worksheets/sheet'));
        
        let sharedStrings: string[] = [];
        let themes: Theme[] = [];
        let style: StyleSheet = {
            fonts: [],
            fills: [],
            borders: [],
            cells: [],
        };
        let workbook: Workbook = {
            workSheets: [],
        };
       
        for(const p of files) {
            // Read workbook.xml
            if (p === 'xl/workbook.xml') {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nWorkbook = parseWorkbookXml(xml);
                    workbook = nWorkbook;
                }
            }
            // Read sharedStrings.xml
            if (p === 'xl/sharedStrings.xml') {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nSharedStrings = parseSharedStringsXml(xml);
                    sharedStrings = nSharedStrings;
                }
            }
            // Read theme1.xml
            if (p === 'xl/theme/theme1.xml' && (options.styles ?? false)) {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nThemes = parseThemeXml(xml);
                    themes = nThemes;
                }
            }
            // Read styles.xml
            if (p === 'xl/styles.xml' && (options.styles ?? false)) {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nStyle = parseStylesXml(xml, themes);
                    style = nStyle;
                }
            }
        }

        for(let i=0; i < sheetFiles.length; i++) {
            // Read sheet[n].xml
            const xml = await nZip.file(sheetFiles[i])?.async('string');
            if (xml) {
                const nSheet = parseWorksheetXml(xml, style, sharedStrings, options.dense ?? false);
                workbook.workSheets[i].dimention = nSheet.dimention;
                workbook.workSheets[i].columnStyles = nSheet.columnStyles;
                workbook.workSheets[i].data = nSheet.data;
                workbook.workSheets[i].mergeCells = nSheet.mergeCells;
            }
        }
        
        return workbook;
    }
}