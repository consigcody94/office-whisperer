/**
 * Excel Generator - Create and manipulate Excel workbooks using ExcelJS
 */
import type { ExcelWorkbookOptions, ExcelFormula, ExcelChart, ExcelPivotTable, ExcelConditionalFormattingRule, ExcelDataValidation, ExcelCellStyle } from '../types.js';
export declare class ExcelGenerator {
    createWorkbook(options: ExcelWorkbookOptions): Promise<Buffer>;
    addPivotTable(filename: string, sheetName: string, pivotTable: ExcelPivotTable): Promise<Buffer>;
    addChart(filename: string, sheetName: string, chart: ExcelChart): Promise<Buffer>;
    addFormulas(filename: string, sheetName: string, formulas: ExcelFormula[]): Promise<Buffer>;
    addConditionalFormatting(filename: string, sheetName: string, range: string, rules: ExcelConditionalFormattingRule[]): Promise<Buffer>;
    addDataValidation(filename: string, sheetName: string, range: string, validation: ExcelDataValidation): Promise<Buffer>;
    freezePanes(filename: string, sheetName: string, row?: number, column?: number): Promise<Buffer>;
    filterSort(filename: string, sheetName: string, range?: string, sortBy?: {
        column: string | number;
        descending?: boolean;
    }[], autoFilter?: boolean): Promise<Buffer>;
    formatCells(filename: string, sheetName: string, range: string, style: ExcelCellStyle): Promise<Buffer>;
    addNamedRange(filename: string, name: string, range: string, sheetName?: string): Promise<Buffer>;
    protectSheet(filename: string, sheetName: string, password?: string, options?: any): Promise<Buffer>;
    mergeWorkbooks(files: string[], outputFilename: string): Promise<Buffer>;
    findReplace(filename: string, find: string, replace: string, sheetName?: string, matchCase?: boolean, matchEntireCell?: boolean, searchFormulas?: boolean): Promise<Buffer>;
    convertToJSON(excelPath: string, sheetName?: string, header?: boolean): Promise<string>;
    convertToCSV(excelPath: string, sheetName?: string): Promise<string>;
    private loadWorkbook;
    private applyRowStyle;
    private normalizeColor;
    private columnToNumber;
}
//# sourceMappingURL=excel-generator.d.ts.map