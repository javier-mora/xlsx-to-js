import { Theme } from "../theme/types";
import { getElementByName, getElementsByName } from "../utils";
import { getColor } from "./getColor";
import { BorderStyleSheet, StyleSheet } from "./types";


export const parseStylesXml = (str: string, themes: Theme[]): StyleSheet => {
    const styleSheet: StyleSheet = {
        fonts: [],
        fills: [],
        borders: [],
        cells: [],
    };
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const styleSheetElement = getElementByName(xmlDoc, 'styleSheet');
    const fontsElement = getElementByName(styleSheetElement, 'fonts');
    const fillsElement = getElementByName(styleSheetElement, 'fills');
    const bordersElement = getElementByName(styleSheetElement, 'borders');
    const cellXfsElement = getElementByName(styleSheetElement, 'cellXfs');
    const fontsArray = getElementsByName(fontsElement, 'font');
    const fillsArray = getElementsByName(fillsElement, 'fill');
    const bordersArray = getElementsByName(bordersElement, 'border');
    const xfsArray = getElementsByName(cellXfsElement, 'xf');

    if (fontsArray) {
        fontsArray.forEach(x => {
            const sz = getElementByName(x, 'sz');
            const color = getElementByName(x, 'color');
            const name = getElementByName(x, 'name');
            const family = getElementByName(x, 'family');
            const bold = getElementByName(x, 'b');
            const italic = getElementByName(x, 'i');

            styleSheet.fonts.push({
                size: +(sz?.getAttribute('val') ?? 0),
                color: getColor(color, themes),
                name: name?.getAttribute('val') ?? '',
                family: +(family?.getAttribute('val') ?? 0),
                bold: bold !== undefined && bold !== null,
                italic: italic !== undefined && italic !== null,
            });
        })
    }

    if (fillsArray) {
        fillsArray.forEach(x => {
            const patternFill = getElementByName(x, 'patternFill');
            const fgColor = getElementByName(patternFill, 'fgColor');
            const bgColor = getElementByName(patternFill, 'bgColor');
            styleSheet.fills.push({
                patternType: patternFill?.getAttribute('patternType') ?? '',
                bgColor: getColor(bgColor, themes),
                fgColor: getColor(fgColor, themes),
            });
        })
    }

    if (bordersArray) {
        bordersArray.forEach(x => {
            const border: BorderStyleSheet = {
                left: { style: '', color: '' },
                right: { style: '', color: '' },
                top: { style: '', color: '' },
                bottom: { style: '', color: '' },
                diagonal: { style: '', color: '' },
            };
            const leftBorder = getElementByName(x, 'left');
            const rightBorder = getElementByName(x, 'right');
            const topBorder = getElementByName(x, 'top');
            const bottomBorder = getElementByName(x, 'bottom');
            const diagonalBorder = getElementByName(x, 'diagonal');

            if (leftBorder) {
                border.left.style = leftBorder.getAttribute('style') ?? '';
                border.left.color = getColor(getElementByName(leftBorder, 'color'), themes);
            }
            if (rightBorder) {
                border.right.style = rightBorder.getAttribute('style') ?? '';
                border.right.color = getColor(getElementByName(rightBorder, 'color'), themes);
            }
            if (topBorder) {
                border.top.style = topBorder.getAttribute('style') ?? '';
                border.top.color = getColor(getElementByName(topBorder, 'color'), themes);
            }
            if (bottomBorder) {
                border.bottom.style = bottomBorder.getAttribute('style') ?? '';
                border.bottom.color = getColor(getElementByName(bottomBorder, 'color'), themes);
            }
            if (diagonalBorder) {
                border.diagonal.style = diagonalBorder.getAttribute('style') ?? '';
                border.diagonal.color = getColor(getElementByName(diagonalBorder, 'color'), themes);
            }

            styleSheet.borders.push(border);
        });
    }

    if (xfsArray) {
        xfsArray.forEach(x => {
            const alignment = getElementByName(x, 'alignment');
            styleSheet.cells.push({
                numFmtId: +(x.getAttribute('numFmtId') ?? 0),
                fontId: +(x.getAttribute('fontId') ?? 0),
                fillId: +(x.getAttribute('fillId') ?? 0),
                borderId: +(x.getAttribute('borderId') ?? 0),
                xfId: +(x.getAttribute('xfId') ?? 0),
                applyFont: +(x.getAttribute('applyFont') ?? 0),
                applyBorder: +(x.getAttribute('applyBorder') ?? 0),
                applyAlignment: +(x.getAttribute('applyAlignment') ?? 0),
                applyFill: +(x.getAttribute('applyFill') ?? 0),
                alignment: alignment
                    ? {
                        horizontal: alignment?.getAttribute('horizontal') ?? '',
                        vertical: alignment?.getAttribute('vertical') ?? 'bottom',
                        wrapText: (alignment?.getAttribute('wrapText') ?? '0') === '1',
                    }
                    : undefined, 
            });
        })
    }

    return styleSheet;
};