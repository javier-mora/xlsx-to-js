import { WorkSheet } from "../worksheet/types";

/** 
 * A collection of worksheets. 
 */
export interface Workbook {
    /** Define a sheet array and its contents. */
    workSheets: WorkSheet[];
}