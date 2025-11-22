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
    addSparklines(filename: string, sheetName: string, dataRange: string, location: string, type: 'line' | 'column' | 'winLoss', options?: any): Promise<Buffer>;
    addArrayFormulas(filename: string, sheetName: string, formulas: Array<{
        cell: string;
        formula: string;
    }>): Promise<Buffer>;
    addSubtotals(filename: string, sheetName: string, range: string, groupBy: number, summaryFunction: 'SUM' | 'COUNT' | 'AVERAGE' | 'MAX' | 'MIN', summaryColumns: number[]): Promise<Buffer>;
    addHyperlinks(filename: string, sheetName: string, links: Array<{
        cell: string;
        url?: string;
        sheet?: string;
        range?: string;
        tooltip?: string;
        displayText?: string;
    }>): Promise<Buffer>;
    addAdvancedChart(filename: string, sheetName: string, chart: {
        type: 'waterfall' | 'funnel' | 'treemap' | 'sunburst' | 'histogram' | 'boxWhisker' | 'pareto';
        title: string;
        dataRange: string;
        position?: {
            row: number;
            col: number;
        };
    }): Promise<Buffer>;
    addSlicers(filename: string, sheetName: string, tableName: string, slicers: Array<{
        columnName: string;
        caption?: string;
        position?: {
            row: number;
            col: number;
        };
    }>): Promise<Buffer>;
    private numberToColumn;
    addPowerQuery(filename: string, sheetName: string, queryName: string, source: any, transformations?: any[]): Promise<Buffer>;
    goalSeek(filename: string, sheetName: string, setCell: string, toValue: number, byChangingCell: string): Promise<Buffer>;
    createDataTable(filename: string, sheetName: string, type: 'oneVariable' | 'twoVariable', formulaCell: string, rowInputCell?: string, columnInputCell?: string, outputRange?: string): Promise<Buffer>;
    manageScenarios(filename: string, sheetName: string, scenarios: Array<{
        name: string;
        changingCells: string[];
        values: (string | number)[];
        comment?: string;
    }>, resultCells?: string[]): Promise<Buffer>;
    createTable(filename: string, sheetName: string, tableName: string, range: string, hasHeaders?: boolean, style?: string, showTotalRow?: boolean): Promise<Buffer>;
    addTableFormula(filename: string, sheetName: string, tableName: string, columnName: string, formula: string): Promise<Buffer>;
    addFormControls(filename: string, sheetName: string, controls: Array<any>): Promise<Buffer>;
    insertImages(filename: string, sheetName: string, images: Array<{
        path: string;
        position: {
            row: number;
            col: number;
        };
        size?: {
            width: number;
            height: number;
        };
        description?: string;
        hyperlink?: string;
    }>): Promise<Buffer>;
    insertShapes(filename: string, sheetName: string, shapes: Array<any>): Promise<Buffer>;
    addSmartArt(filename: string, sheetName: string, smartArt: any): Promise<Buffer>;
    configurePageSetup(filename: string, sheetName: string, pageSetup: any): Promise<Buffer>;
    setHeaderFooter(filename: string, sheetName: string, header?: any, footer?: any, differentFirstPage?: boolean, differentOddEven?: boolean): Promise<Buffer>;
    addPageBreaks(filename: string, sheetName: string, horizontalBreaks?: number[], verticalBreaks?: number[]): Promise<Buffer>;
    enableTrackChanges(filename: string, enable: boolean, highlightChanges?: boolean): Promise<Buffer>;
    shareWorkbook(filename: string, share: boolean, allowChanges?: boolean, password?: string): Promise<Buffer>;
    addComments(filename: string, sheetName: string, comments: Array<{
        cell: string;
        text: string;
        author?: string;
        visible?: boolean;
    }>): Promise<Buffer>;
    private getTableColumnsFromRange;
    private getTableRowsFromRange;
    private getPaperSizeCode;
    private formatHeaderFooter;
}
//# sourceMappingURL=excel-generator.d.ts.map