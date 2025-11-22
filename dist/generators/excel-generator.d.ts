/**
 * Excel Generator - Create and manipulate Excel workbooks using ExcelJS
 */
import type { ExcelWorkbookOptions, ExcelFormula, ExcelChart } from '../types.js';
export declare class ExcelGenerator {
    createWorkbook(options: ExcelWorkbookOptions): Promise<Buffer>;
    addFormulas(filename: string, sheetName: string, formulas: ExcelFormula[]): Promise<Buffer>;
    addChart(filename: string, sheetName: string, chart: ExcelChart): Promise<Buffer>;
    private applyRowStyle;
    convertToCSV(excelPath: string, sheetName?: string): Promise<string>;
}
//# sourceMappingURL=excel-generator.d.ts.map