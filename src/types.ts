/**
 * Office Whisperer - Type Definitions
 * Comprehensive TypeScript types for Microsoft Office Suite automation via MCP
 */

// ============================================================================
// MCP Protocol Types
// ============================================================================

export interface MCPRequest {
  jsonrpc: '2.0';
  id?: string | number;
  method: string;
  params?: Record<string, unknown>;
}

export interface MCPResponse {
  jsonrpc: '2.0';
  id: string | number;
  result?: unknown;
  error?: MCPError;
}

export interface MCPError {
  code: number;
  message: string;
  data?: unknown;
}

export interface MCPTool {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

// ============================================================================
// Excel Types
// ============================================================================

export interface ExcelWorkbookOptions {
  filename: string;
  sheets: ExcelSheet[];
}

export interface ExcelSheet {
  name: string;
  data?: (string | number | boolean | null)[][];
  columns?: ExcelColumn[];
  rows?: ExcelRow[];
  charts?: ExcelChart[];
  pivotTables?: ExcelPivotTable[];
}

export interface ExcelColumn {
  header: string;
  key: string;
  width?: number;
  style?: ExcelCellStyle;
}

export interface ExcelRow {
  values: (string | number | boolean | null)[];
  style?: ExcelCellStyle;
}

export interface ExcelCellStyle {
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
  };
  fill?: {
    type: 'pattern';
    pattern: string;
    fgColor: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  border?: {
    top?: ExcelBorder;
    bottom?: ExcelBorder;
    left?: ExcelBorder;
    right?: ExcelBorder;
  };
  numFmt?: string;
}

export interface ExcelBorder {
  style: 'thin' | 'medium' | 'thick' | 'double';
  color: string;
}

export interface ExcelChart {
  type: 'line' | 'bar' | 'pie' | 'scatter' | 'area';
  title: string;
  dataRange: string;
  position?: {
    row: number;
    col: number;
  };
}

export interface ExcelPivotTable {
  name: string;
  dataRange: string;
  rows: string[];
  columns: string[];
  values: string[];
  filters?: string[];
}

export interface ExcelFormula {
  cell: string;
  formula: string;
}

// ============================================================================
// Word Types
// ============================================================================

export interface WordDocumentOptions {
  filename: string;
  sections: WordSection[];
  styles?: WordStyles;
}

export interface WordSection {
  properties?: {
    page?: {
      size?: { width: number; height: number };
      margin?: { top: number; right: number; bottom: number; left: number };
    };
  };
  headers?: WordHeaderFooter[];
  footers?: WordHeaderFooter[];
  children: WordElement[];
}

export interface WordHeaderFooter {
  type: 'default' | 'first' | 'even';
  children: WordElement[];
}

export type WordElement =
  | WordParagraph
  | WordTable
  | WordImage
  | WordTOC
  | WordPageBreak;

export interface WordParagraph {
  type: 'paragraph';
  text?: string;
  children?: WordRun[];
  heading?: 'Heading1' | 'Heading2' | 'Heading3' | 'Heading4' | 'Heading5' | 'Heading6';
  alignment?: 'left' | 'center' | 'right' | 'justified';
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
  };
  bullet?: {
    level: number;
  };
  numbering?: {
    level: number;
    reference: string;
  };
}

export interface WordRun {
  text: string;
  bold?: boolean;
  italics?: boolean;
  underline?: { type: 'single' | 'double' | 'thick' | 'dotted' };
  strike?: boolean;
  color?: string;
  size?: number;
  font?: string;
  highlight?: string;
}

export interface WordTable {
  type: 'table';
  rows: WordTableRow[];
  width?: {
    size: number;
    type: 'dxa' | 'pct' | 'auto';
  };
  borders?: WordTableBorders;
}

export interface WordTableRow {
  cells: WordTableCell[];
  height?: number;
  cantSplit?: boolean;
  tableHeader?: boolean;
}

export interface WordTableCell {
  children: WordParagraph[];
  shading?: {
    fill: string;
    color?: string;
  };
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  columnSpan?: number;
  rowSpan?: number;
}

export interface WordTableBorders {
  top?: { style: string; size: number; color: string };
  bottom?: { style: string; size: number; color: string };
  left?: { style: string; size: number; color: string };
  right?: { style: string; size: number; color: string };
  insideHorizontal?: { style: string; size: number; color: string };
  insideVertical?: { style: string; size: number; color: string };
}

export interface WordImage {
  type: 'image';
  path: string;
  transformation?: {
    width: number;
    height: number;
  };
}

export interface WordTOC {
  type: 'toc';
  title?: string;
}

export interface WordPageBreak {
  type: 'pageBreak';
}

export interface WordStyles {
  default?: {
    document?: {
      run?: {
        font?: string;
        size?: number;
      };
      paragraph?: {
        spacing?: {
          line?: number;
        };
      };
    };
  };
  paragraphStyles?: Array<{
    id: string;
    name: string;
    basedOn?: string;
    next?: string;
    run?: {
      font?: string;
      size?: number;
      bold?: boolean;
      italics?: boolean;
      color?: string;
    };
    paragraph?: {
      spacing?: {
        before?: number;
        after?: number;
        line?: number;
      };
      alignment?: 'left' | 'center' | 'right' | 'justified';
    };
  }>;
}

// ============================================================================
// PowerPoint Types
// ============================================================================

export interface PowerPointPresentationOptions {
  filename: string;
  title?: string;
  author?: string;
  company?: string;
  theme?: 'default' | 'light' | 'dark' | 'colorful';
  slides: PowerPointSlide[];
}

export interface PowerPointSlide {
  layout: 'title' | 'content' | 'section' | 'comparison' | 'blank';
  title?: string;
  subtitle?: string;
  content?: PowerPointContent[];
  notes?: string;
  backgroundColor?: string;
  backgroundImage?: string;
}

export type PowerPointContent =
  | PowerPointText
  | PowerPointImage
  | PowerPointShape
  | PowerPointTable
  | PowerPointChart;

export interface PowerPointText {
  type: 'text';
  text: string;
  x: number | string;
  y: number | string;
  w?: number | string;
  h?: number | string;
  fontSize?: number;
  fontFace?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right' | 'justify';
  valign?: 'top' | 'middle' | 'bottom';
  bullet?: boolean | { type: string; code?: string };
}

export interface PowerPointImage {
  type: 'image';
  path: string;
  x: number | string;
  y: number | string;
  w: number | string;
  h: number | string;
  sizing?: {
    type: 'cover' | 'contain' | 'crop';
  };
}

export interface PowerPointShape {
  type: 'shape';
  shape: 'rectangle' | 'ellipse' | 'triangle' | 'line' | 'arrow';
  x: number | string;
  y: number | string;
  w: number | string;
  h: number | string;
  fill?: { color: string };
  line?: { color: string; width: number };
}

export interface PowerPointTable {
  type: 'table';
  x: number | string;
  y: number | string;
  w?: number | string;
  rows: (string | number)[][];
  colW?: number[];
  rowH?: number;
  fontSize?: number;
  color?: string;
  fill?: string;
  border?: { type: string; color: string; pt: number };
}

export interface PowerPointChart {
  type: 'chart';
  chartType: 'bar' | 'line' | 'pie' | 'scatter' | 'area';
  x: number | string;
  y: number | string;
  w: number | string;
  h: number | string;
  title?: string;
  data: PowerPointChartData[];
}

export interface PowerPointChartData {
  name: string;
  labels: string[];
  values: number[];
}

// ============================================================================
// MCP Tool Arguments
// ============================================================================

export interface CreateExcelArgs {
  filename: string;
  sheets: ExcelSheet[];
  outputPath?: string;
}

export interface CreateWordArgs {
  filename: string;
  title?: string;
  sections: WordSection[];
  outputPath?: string;
}

export interface CreatePowerPointArgs {
  filename: string;
  title?: string;
  theme?: 'default' | 'light' | 'dark' | 'colorful';
  slides: PowerPointSlide[];
  outputPath?: string;
}

export interface AddExcelFormulaArgs {
  filename: string;
  sheetName: string;
  formulas: ExcelFormula[];
}

export interface AddExcelChartArgs {
  filename: string;
  sheetName: string;
  chart: ExcelChart;
}

export interface FormatExcelCellsArgs {
  filename: string;
  sheetName: string;
  range: string;
  style: ExcelCellStyle;
}

export interface ConvertExcelToCSVArgs {
  excelPath: string;
  sheetName?: string;
  outputPath?: string;
}

export interface MergeWordDocumentsArgs {
  documents: string[];
  outputPath: string;
}

export interface AddPowerPointSlideArgs {
  filename: string;
  slide: PowerPointSlide;
  position?: number;
}
