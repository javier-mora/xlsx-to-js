import { getColumnIndex } from '../utils/utils'

/**
 * Gets the rows and columns spanned by a range.
 * @param range Range in *A1:B2* format.
 * @returns 
 */
export function getRangeDetails(range: string): { origin: string, colspan: number, rowspan: number } {
    const [start, end] = range.split(':');
  
    const startColumn = start.match(/[A-Z]+/)![0];
    const startRow = parseInt(start.match(/\d+/)![0]);
  
    const endColumn = end.match(/[A-Z]+/)![0];
    const endRow = parseInt(end.match(/\d+/)![0]);
  
    const startColIndex = getColumnIndex(startColumn);
    const endColIndex = getColumnIndex(endColumn);
  
    const colspan = (endColIndex - startColIndex) + 1;
    const rowspan = (endRow - startRow) + 1;
  
    return {
        origin: start,
        colspan: colspan,
        rowspan: rowspan
    };
  }