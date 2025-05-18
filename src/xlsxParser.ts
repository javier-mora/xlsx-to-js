import JSZip from "jszip";
import { XlsxParserOptions } from "./types";
import { parseThemeXml, Theme } from "./core/theme";
import { parseStylesXml, StyleSheet } from "./core/style";
import { parseWorkbookXml, Workbook } from "./core/workbook";
import { parseSharedStringsXml } from "./core/sharedString";
import { parseWorksheetXml } from "./core/worksheet";
import { DrawingFile, MediaFile } from "./core/drawing/types";
import { parseDrawingXml } from "./core/drawing";

export class XlsxParser {
    
    /**
     * Extract data from spreadsheet bytes stored in a `ArrayBuffer`.
     * @param file Workbook file
     * @param options Parser options
     * @returns 
     */
    async readFile (file: ArrayBuffer, options: XlsxParserOptions = { dense: false, styles: false, drawings: false, skipHiddenRows: false }) {
        const nZip = new JSZip();
        const result = await nZip.loadAsync(file);
        const files: string[] = [];
        result.forEach((p) => files.push(p));
        const sheetFiles = files.filter(x => x.includes('xl/worksheets/sheet'));
        const drawingFiles = files.filter(x => x.includes('xl/drawings/drawing'));
        const drawingRelFiles = files.filter(x => x.includes('xl/drawings/_rels/drawing'));
        const mediaFiles = files.filter(x => x.includes('xl/media/'));
        
        let sharedStrings: string[] = [];
        let themes: Theme[] = [];
        let drawings: DrawingFile[] = [];
        const media: MediaFile[] = [];
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

        for(let i=0; i < mediaFiles.length; i++) {
            // Read image[n].png
            const base64 = (await nZip.file(mediaFiles[i])?.async('base64')) ?? '';
            media.push({
                name: mediaFiles[i].split('/').pop() || '',
                base64: base64,
            });
        }

        for(let i=0; i < drawingFiles.length; i++) {
            // Read drawing[n].xml
            const xml = await nZip.file(drawingFiles[i])?.async('string');
            // Read _rels/drawing[n].xml.rels
            const relName = drawingFiles[i].split('/').pop() || '';
            const relFilePath = drawingRelFiles.find(x => x.includes(`${relName}.rels`)) ?? '';
            const rel = await nZip.file(relFilePath)?.async('string');

            if (xml && rel) {
                const nDrawing = parseDrawingXml(xml, rel, media);
                drawings.push({ src: `drawing${i+1}.xml`, drawings: nDrawing});
            }
        }

        for(let i=0; i < sheetFiles.length; i++) {
            // Read sheet[n].xml
            const xml = await nZip.file(sheetFiles[i])?.async('string');
            // Read _rels/sheet[n].xml.rels
            const sheetName = sheetFiles[i].split('/').pop() || '';
            const relFilePath = `xl/worksheets/_rels/${sheetName}.rels`;
            const rel = await nZip.file(relFilePath)?.async('string');
            if (xml) {
                const nSheet = parseWorksheetXml(xml, rel, style, sharedStrings, drawings, options.dense ?? false, options.skipHiddenRows ?? false);
                workbook.workSheets[i].dimention = nSheet.dimention;
                workbook.workSheets[i].columnStyles = nSheet.columnStyles;
                workbook.workSheets[i].rowStyles = nSheet.rowStyles;
                workbook.workSheets[i].data = nSheet.data;
                workbook.workSheets[i].mergeCells = nSheet.mergeCells;
                workbook.workSheets[i].drawings = nSheet.drawings;
            }
        }
        
        return workbook;
    }
}