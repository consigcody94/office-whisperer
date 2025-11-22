/**
 * Office Whisperer - Type Definitions
 * Comprehensive TypeScript types for Microsoft Office Suite automation via MCP
 */
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
export interface WordDocumentOptions {
    filename: string;
    sections: WordSection[];
    styles?: WordStyles;
}
export interface WordSection {
    properties?: {
        page?: {
            size?: {
                width: number;
                height: number;
            };
            margin?: {
                top: number;
                right: number;
                bottom: number;
                left: number;
            };
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
export type WordElement = WordParagraph | WordTable | WordImage | WordTOC | WordPageBreak;
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
    underline?: {
        type: 'single' | 'double' | 'thick' | 'dotted';
    };
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
    top?: {
        style: string;
        size: number;
        color: string;
    };
    bottom?: {
        style: string;
        size: number;
        color: string;
    };
    left?: {
        style: string;
        size: number;
        color: string;
    };
    right?: {
        style: string;
        size: number;
        color: string;
    };
    insideHorizontal?: {
        style: string;
        size: number;
        color: string;
    };
    insideVertical?: {
        style: string;
        size: number;
        color: string;
    };
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
export type PowerPointContent = PowerPointText | PowerPointImage | PowerPointShape | PowerPointTable | PowerPointChart;
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
    bullet?: boolean | {
        type: string;
        code?: string;
    };
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
    fill?: {
        color: string;
    };
    line?: {
        color: string;
        width: number;
    };
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
    border?: {
        type: string;
        color: string;
        pt: number;
    };
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
export interface ExcelAddPivotTableArgs {
    filename: string;
    sheetName: string;
    pivotTable: ExcelPivotTable;
    outputPath?: string;
}
export interface ExcelAddChartArgs {
    filename: string;
    sheetName: string;
    chart: ExcelChart;
    outputPath?: string;
}
export interface ExcelAddFormulaArgs {
    filename: string;
    sheetName: string;
    formulas: ExcelFormula[];
    outputPath?: string;
}
export interface ExcelConditionalFormattingArgs {
    filename: string;
    sheetName: string;
    range: string;
    rules: ExcelConditionalFormattingRule[];
    outputPath?: string;
}
export interface ExcelConditionalFormattingRule {
    type: 'colorScale' | 'dataBar' | 'iconSet' | 'formulaBased' | 'cellValue';
    formula?: string;
    priority?: number;
    color?: string;
    gradient?: {
        start: string;
        middle?: string;
        end: string;
    };
    iconSet?: 'ThreeArrows' | 'ThreeFlags' | 'FourRating' | 'FiveQuarters';
    operator?: 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'notEqual';
    values?: (number | string)[];
}
export interface ExcelDataValidationArgs {
    filename: string;
    sheetName: string;
    range: string;
    validation: ExcelDataValidation;
    outputPath?: string;
}
export interface ExcelDataValidation {
    type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
    formula?: string;
    values?: string[];
    operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan';
    min?: number | string;
    max?: number | string;
    allowBlank?: boolean;
    showErrorMessage?: boolean;
    errorTitle?: string;
    error?: string;
    showInputMessage?: boolean;
    promptTitle?: string;
    prompt?: string;
}
export interface ExcelFreezePanesArgs {
    filename: string;
    sheetName: string;
    row?: number;
    column?: number;
    outputPath?: string;
}
export interface ExcelFilterSortArgs {
    filename: string;
    sheetName: string;
    range?: string;
    sortBy?: {
        column: string | number;
        descending?: boolean;
    }[];
    autoFilter?: boolean;
    outputPath?: string;
}
export interface ExcelFormatCellsArgs {
    filename: string;
    sheetName: string;
    range: string;
    style: ExcelCellStyle;
    outputPath?: string;
}
export interface ExcelNamedRangeArgs {
    filename: string;
    name: string;
    range: string;
    sheetName?: string;
    outputPath?: string;
}
export interface ExcelProtectSheetArgs {
    filename: string;
    sheetName: string;
    password?: string;
    options?: {
        selectLockedCells?: boolean;
        selectUnlockedCells?: boolean;
        formatCells?: boolean;
        formatColumns?: boolean;
        formatRows?: boolean;
        insertColumns?: boolean;
        insertRows?: boolean;
        insertHyperlinks?: boolean;
        deleteColumns?: boolean;
        deleteRows?: boolean;
        sort?: boolean;
        autoFilter?: boolean;
        pivotTables?: boolean;
    };
    outputPath?: string;
}
export interface ExcelMergeWorkbooksArgs {
    files: string[];
    outputFilename: string;
    outputPath?: string;
}
export interface ExcelFindReplaceArgs {
    filename: string;
    sheetName?: string;
    find: string;
    replace: string;
    matchCase?: boolean;
    matchEntireCell?: boolean;
    searchFormulas?: boolean;
    outputPath?: string;
}
export interface ExcelToJSONArgs {
    excelPath: string;
    sheetName?: string;
    outputPath?: string;
    header?: boolean;
}
export interface WordAddTOCArgs {
    filename: string;
    outputPath?: string;
    title?: string;
    hyperlinks?: boolean;
    levels?: number;
}
export interface WordMailMergeArgs {
    templatePath: string;
    dataSource: Record<string, string | number>[];
    outputPath?: string;
    outputFilename?: string;
}
export interface WordFindReplaceArgs {
    filename: string;
    find: string;
    replace: string;
    matchCase?: boolean;
    matchWholeWord?: boolean;
    formatting?: {
        bold?: boolean;
        italic?: boolean;
        color?: string;
    };
    outputPath?: string;
}
export interface WordAddCommentArgs {
    filename: string;
    text: string;
    comment: string;
    author?: string;
    outputPath?: string;
}
export interface WordFormatStylesArgs {
    filename: string;
    styles: WordStyles;
    outputPath?: string;
}
export interface WordInsertImageArgs {
    filename: string;
    imagePath: string;
    position?: {
        x?: number;
        y?: number;
    };
    size?: {
        width?: number;
        height?: number;
    };
    wrapping?: 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';
    outputPath?: string;
}
export interface WordAddHeaderFooterArgs {
    filename: string;
    type: 'header' | 'footer';
    content: WordElement[];
    sectionType?: 'default' | 'first' | 'even';
    outputPath?: string;
}
export interface WordCompareDocumentsArgs {
    originalPath: string;
    revisedPath: string;
    outputPath?: string;
    author?: string;
}
export interface WordToPDFArgs {
    filename: string;
    outputPath?: string;
}
export interface PPTAddTransitionArgs {
    filename: string;
    slideNumber?: number;
    transition: PPTTransition;
    outputPath?: string;
}
export interface PPTTransition {
    type: 'fade' | 'push' | 'wipe' | 'split' | 'reveal' | 'randomBars' | 'circle' | 'dissolve';
    duration?: number;
    direction?: 'left' | 'right' | 'up' | 'down';
}
export interface PPTAddAnimationArgs {
    filename: string;
    slideNumber: number;
    objectId?: string;
    animation: PPTAnimation;
    outputPath?: string;
}
export interface PPTAnimation {
    type: 'entrance' | 'emphasis' | 'exit' | 'motion';
    effect: 'appear' | 'fade' | 'fly' | 'float' | 'split' | 'wipe' | 'shape' | 'wheel' | 'randomBars' | 'grow' | 'zoom' | 'swivel' | 'bounce';
    duration?: number;
    delay?: number;
    direction?: 'left' | 'right' | 'up' | 'down';
}
export interface PPTAddNotesArgs {
    filename: string;
    slideNumber: number;
    notes: string;
    outputPath?: string;
}
export interface PPTDuplicateSlideArgs {
    filename: string;
    slideNumber: number;
    position?: number;
    outputPath?: string;
}
export interface PPTReorderSlidesArgs {
    filename: string;
    slideOrder: number[];
    outputPath?: string;
}
export interface PPTExportPDFArgs {
    filename: string;
    outputPath?: string;
}
export interface PPTAddMediaArgs {
    filename: string;
    slideNumber: number;
    mediaPath: string;
    mediaType: 'video' | 'audio';
    position?: {
        x: number;
        y: number;
    };
    size?: {
        width: number;
        height: number;
    };
    outputPath?: string;
}
export interface OutlookSendEmailArgs {
    to: string | string[];
    subject: string;
    body: string;
    cc?: string | string[];
    bcc?: string | string[];
    attachments?: OutlookAttachment[];
    html?: boolean;
    priority?: 'high' | 'normal' | 'low';
    smtpConfig?: OutlookSMTPConfig;
}
export interface OutlookAttachment {
    filename: string;
    path?: string;
    content?: string | Buffer;
}
export interface OutlookSMTPConfig {
    host: string;
    port: number;
    secure?: boolean;
    auth?: {
        user: string;
        pass: string;
    };
}
export interface OutlookCreateMeetingArgs {
    subject: string;
    startTime: string;
    endTime: string;
    location?: string;
    attendees?: OutlookAttendee[];
    description?: string;
    reminder?: number;
    outputPath?: string;
}
export interface OutlookAttendee {
    email: string;
    name?: string;
    required?: boolean;
}
export interface OutlookAddContactArgs {
    firstName: string;
    lastName: string;
    email?: string;
    phone?: string;
    company?: string;
    jobTitle?: string;
    address?: string;
    outputPath?: string;
}
export interface OutlookCreateTaskArgs {
    subject: string;
    dueDate?: string;
    priority?: 'high' | 'normal' | 'low';
    status?: 'notStarted' | 'inProgress' | 'completed' | 'waiting' | 'deferred';
    category?: string;
    reminder?: string;
    notes?: string;
    outputPath?: string;
}
export interface OutlookSetRuleArgs {
    name: string;
    conditions: OutlookRuleCondition[];
    actions: OutlookRuleAction[];
    outputPath?: string;
}
export interface OutlookRuleCondition {
    type: 'from' | 'subject' | 'body' | 'recipient' | 'attachment';
    value: string;
    operator?: 'contains' | 'equals' | 'startsWith' | 'endsWith';
}
export interface OutlookRuleAction {
    type: 'move' | 'copy' | 'delete' | 'forward' | 'flag' | 'category';
    value: string;
}
export interface ExcelAddSparklinesArgs {
    filename: string;
    sheetName: string;
    dataRange: string;
    location: string;
    type: 'line' | 'column' | 'winLoss';
    options?: {
        lineWeight?: number;
        markers?: boolean;
        high?: boolean;
        low?: boolean;
        first?: boolean;
        last?: boolean;
        negative?: boolean;
        displayEmptyCellsAs?: 'gaps' | 'zero' | 'connect';
    };
    outputPath?: string;
}
export interface ExcelArrayFormulasArgs {
    filename: string;
    sheetName: string;
    formulas: Array<{
        cell: string;
        formula: string;
    }>;
    outputPath?: string;
}
export interface ExcelAddSubtotalsArgs {
    filename: string;
    sheetName: string;
    range: string;
    groupBy: number;
    summaryFunction: 'SUM' | 'COUNT' | 'AVERAGE' | 'MAX' | 'MIN';
    summaryColumns: number[];
    replaceExisting?: boolean;
    pageBreakBetweenGroups?: boolean;
    summaryBelowData?: boolean;
    outputPath?: string;
}
export interface ExcelAddHyperlinksArgs {
    filename: string;
    sheetName: string;
    links: Array<{
        cell: string;
        url?: string;
        sheet?: string;
        range?: string;
        tooltip?: string;
        displayText?: string;
    }>;
    outputPath?: string;
}
export interface ExcelAdvancedChartsArgs {
    filename: string;
    sheetName: string;
    chart: {
        type: 'waterfall' | 'funnel' | 'treemap' | 'sunburst' | 'histogram' | 'boxWhisker' | 'pareto';
        title: string;
        dataRange: string;
        categories?: string;
        values?: string;
        position?: {
            row: number;
            col: number;
        };
    };
    outputPath?: string;
}
export interface ExcelAddSlicersArgs {
    filename: string;
    sheetName: string;
    tableName: string;
    slicers: Array<{
        columnName: string;
        caption?: string;
        position?: {
            row: number;
            col: number;
        };
        style?: string;
    }>;
    outputPath?: string;
}
export interface WordTrackChangesArgs {
    filename: string;
    enable: boolean;
    author?: string;
    showMarkup?: boolean;
    trackFormatting?: boolean;
    trackMoves?: boolean;
    outputPath?: string;
}
export interface WordAddFootnotesArgs {
    filename: string;
    footnotes: Array<{
        text: string;
        note: string;
        type?: 'footnote' | 'endnote';
    }>;
    outputPath?: string;
}
export interface WordAddBookmarksArgs {
    filename: string;
    bookmarks: Array<{
        name: string;
        text: string;
    }>;
    outputPath?: string;
}
export interface WordAddSectionBreaksArgs {
    filename: string;
    breaks: Array<{
        position: number;
        type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
    }>;
    outputPath?: string;
}
export interface WordAddTextBoxesArgs {
    filename: string;
    textBoxes: Array<{
        text: string;
        position?: {
            x: number;
            y: number;
        };
        width?: number;
        height?: number;
        wrapping?: 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom';
        border?: boolean;
        fill?: string;
    }>;
    outputPath?: string;
}
export interface WordAddCrossReferencesArgs {
    filename: string;
    references: Array<{
        bookmarkName: string;
        referenceType: 'pageNumber' | 'text' | 'above/below';
        insertText?: string;
    }>;
    outputPath?: string;
}
export interface PPTMasterSlidesArgs {
    filename: string;
    masterSlide: {
        name: string;
        background?: {
            color?: string;
            image?: string;
        };
        placeholders?: Array<{
            type: 'title' | 'body' | 'footer' | 'slideNumber' | 'date';
            x: number;
            y: number;
            w: number;
            h: number;
        }>;
        fonts?: {
            title?: string;
            body?: string;
        };
        colors?: {
            accent1?: string;
            accent2?: string;
            accent3?: string;
        };
    };
    outputPath?: string;
}
export interface PPTAddHyperlinksArgs {
    filename: string;
    slideNumber: number;
    links: Array<{
        text: string;
        url?: string;
        slide?: number;
        tooltip?: string;
    }>;
    outputPath?: string;
}
export interface PPTAddSectionsArgs {
    filename: string;
    sections: Array<{
        name: string;
        startSlide: number;
    }>;
    outputPath?: string;
}
export interface PPTMorphTransitionArgs {
    filename: string;
    fromSlide: number;
    toSlide: number;
    duration?: number;
    outputPath?: string;
}
export interface PPTAddActionButtonsArgs {
    filename: string;
    slideNumber: number;
    buttons: Array<{
        text: string;
        action: 'nextSlide' | 'previousSlide' | 'firstSlide' | 'lastSlide' | 'endShow' | 'customSlide';
        targetSlide?: number;
        x: number;
        y: number;
        w?: number;
        h?: number;
    }>;
    outputPath?: string;
}
export interface OutlookReadEmailsArgs {
    folder?: string;
    limit?: number;
    unreadOnly?: boolean;
    since?: string;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
export interface OutlookSearchEmailsArgs {
    query: string;
    searchIn?: ('subject' | 'from' | 'body' | 'to')[];
    folder?: string;
    limit?: number;
    since?: string;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
export interface OutlookRecurringMeetingArgs {
    subject: string;
    startTime: string;
    endTime: string;
    recurrence: {
        frequency: 'daily' | 'weekly' | 'monthly' | 'yearly';
        interval?: number;
        daysOfWeek?: ('MO' | 'TU' | 'WE' | 'TH' | 'FR' | 'SA' | 'SU')[];
        until?: string;
        count?: number;
    };
    location?: string;
    attendees?: OutlookAttendee[];
    description?: string;
    outputPath?: string;
}
export interface OutlookEmailTemplateArgs {
    name: string;
    subject: string;
    body: string;
    html?: boolean;
    placeholders?: string[];
    outputPath?: string;
}
export interface OutlookMarkReadArgs {
    messageIds: string[];
    markAsRead: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
export interface OutlookArchiveEmailArgs {
    messageIds: string[];
    archiveFolder?: string;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
export interface OutlookCalendarViewArgs {
    startDate: string;
    endDate: string;
    viewType: 'day' | 'week' | 'month' | 'agenda';
    outputFormat?: 'ics' | 'json';
    outputPath?: string;
}
export interface OutlookSearchContactsArgs {
    query: string;
    searchIn?: ('name' | 'email' | 'company' | 'phone')[];
    outputFormat?: 'vcf' | 'json';
    outputPath?: string;
}
//# sourceMappingURL=types.d.ts.map