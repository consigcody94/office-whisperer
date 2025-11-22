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
export interface ExcelPowerQueryArgs {
    filename: string;
    sheetName: string;
    queryName: string;
    source: {
        type: 'table' | 'range' | 'csv' | 'json' | 'web';
        location: string;
    };
    transformations?: Array<{
        step: 'filter' | 'sort' | 'groupBy' | 'pivot' | 'unpivot' | 'merge' | 'append' | 'removeColumns' | 'renameColumns' | 'changeType' | 'fillDown' | 'replaceValues';
        column?: string;
        condition?: string;
        value?: any;
        columns?: string[];
        newName?: string;
        dataType?: 'text' | 'number' | 'date' | 'boolean';
    }>;
    loadTo?: 'table' | 'pivotTable' | 'connection';
    outputPath?: string;
}
export interface ExcelGoalSeekArgs {
    filename: string;
    sheetName: string;
    setCell: string;
    toValue: number;
    byChangingCell: string;
    outputPath?: string;
}
export interface ExcelDataTableArgs {
    filename: string;
    sheetName: string;
    type: 'oneVariable' | 'twoVariable';
    formulaCell: string;
    rowInputCell?: string;
    columnInputCell?: string;
    outputRange: string;
    outputPath?: string;
}
export interface ExcelScenarioManagerArgs {
    filename: string;
    sheetName: string;
    scenarios: Array<{
        name: string;
        changingCells: string[];
        values: (string | number)[];
        comment?: string;
    }>;
    resultCells?: string[];
    outputPath?: string;
}
export interface ExcelCreateTableArgs {
    filename: string;
    sheetName: string;
    tableName: string;
    range: string;
    hasHeaders?: boolean;
    style?: 'TableStyleLight1' | 'TableStyleLight2' | 'TableStyleMedium1' | 'TableStyleMedium2' | 'TableStyleDark1' | 'TableStyleDark2';
    showTotalRow?: boolean;
    showFilterButton?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
    outputPath?: string;
}
export interface ExcelTableFormulaArgs {
    filename: string;
    sheetName: string;
    tableName: string;
    columnName: string;
    formula: string;
    outputPath?: string;
}
export interface ExcelFormControlArgs {
    filename: string;
    sheetName: string;
    controls: Array<{
        type: 'button' | 'checkbox' | 'comboBox' | 'listBox' | 'spinner' | 'scrollBar' | 'optionButton' | 'groupBox';
        name: string;
        position: {
            row: number;
            col: number;
        };
        size?: {
            width: number;
            height: number;
        };
        linkedCell?: string;
        inputRange?: string;
        min?: number;
        max?: number;
        increment?: number;
        label?: string;
        macro?: string;
    }>;
    outputPath?: string;
}
export interface ExcelInsertImageArgs {
    filename: string;
    sheetName: string;
    images: Array<{
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
    }>;
    outputPath?: string;
}
export interface ExcelInsertShapeArgs {
    filename: string;
    sheetName: string;
    shapes: Array<{
        type: 'rectangle' | 'roundedRectangle' | 'ellipse' | 'triangle' | 'diamond' | 'pentagon' | 'hexagon' | 'star' | 'arrow' | 'line' | 'connector';
        position: {
            row: number;
            col: number;
        };
        size: {
            width: number;
            height: number;
        };
        fill?: {
            color: string;
            transparency?: number;
        };
        border?: {
            color: string;
            width: number;
            style?: 'solid' | 'dash' | 'dot';
        };
        text?: string;
        rotation?: number;
    }>;
    outputPath?: string;
}
export interface ExcelSmartArtArgs {
    filename: string;
    sheetName: string;
    smartArt: {
        type: 'list' | 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid';
        layout: string;
        position: {
            row: number;
            col: number;
        };
        size?: {
            width: number;
            height: number;
        };
        items: Array<{
            text: string;
            level?: number;
            children?: string[];
        }>;
        colorScheme?: 'colorful' | 'accent1' | 'accent2' | 'accent3';
    };
    outputPath?: string;
}
export interface ExcelPageSetupArgs {
    filename: string;
    sheetName: string;
    pageSetup: {
        orientation?: 'portrait' | 'landscape';
        paperSize?: 'letter' | 'legal' | 'A4' | 'A3' | 'tabloid';
        fitToPage?: {
            width: number;
            height: number;
        };
        scale?: number;
        margins?: {
            top?: number;
            bottom?: number;
            left?: number;
            right?: number;
            header?: number;
            footer?: number;
        };
        centerHorizontally?: boolean;
        centerVertically?: boolean;
        printGridlines?: boolean;
        printHeadings?: boolean;
        blackAndWhite?: boolean;
        draftQuality?: boolean;
        printComments?: 'asDisplayed' | 'atEnd' | 'none';
        printErrors?: 'displayed' | 'blank' | 'dash' | 'NA';
        printArea?: string;
        printTitles?: {
            rows?: string;
            columns?: string;
        };
    };
    outputPath?: string;
}
export interface ExcelHeaderFooterArgs {
    filename: string;
    sheetName: string;
    header?: {
        left?: string;
        center?: string;
        right?: string;
    };
    footer?: {
        left?: string;
        center?: string;
        right?: string;
    };
    differentFirstPage?: boolean;
    differentOddEven?: boolean;
    outputPath?: string;
}
export interface ExcelPageBreaksArgs {
    filename: string;
    sheetName: string;
    horizontalBreaks?: number[];
    verticalBreaks?: number[];
    outputPath?: string;
}
export interface ExcelTrackChangesArgs {
    filename: string;
    enable: boolean;
    highlightChanges?: boolean;
    listChangesOnNewSheet?: boolean;
    trackWhile?: 'editing' | 'shared';
    outputPath?: string;
}
export interface ExcelShareWorkbookArgs {
    filename: string;
    share: boolean;
    allowChanges?: boolean;
    protectSharing?: boolean;
    password?: string;
    updateInterval?: number;
    outputPath?: string;
}
export interface ExcelCompareVersionsArgs {
    originalFile: string;
    revisedFile: string;
    outputPath?: string;
    highlightDifferences?: boolean;
}
export interface ExcelCommentArgs {
    filename: string;
    sheetName: string;
    comments: Array<{
        cell: string;
        text: string;
        author?: string;
        visible?: boolean;
    }>;
    outputPath?: string;
}
export interface WordBibliographyArgs {
    filename: string;
    sources: Array<{
        type: 'book' | 'article' | 'website' | 'journal' | 'report' | 'conference';
        author?: string;
        title: string;
        year?: number;
        publisher?: string;
        city?: string;
        url?: string;
        pages?: string;
        volume?: string;
        issue?: string;
        tag?: string;
    }>;
    style?: 'APA' | 'MLA' | 'Chicago' | 'Harvard' | 'IEEE';
    insertAt?: 'end' | 'cursor';
    outputPath?: string;
}
export interface WordCitationArgs {
    filename: string;
    citations: Array<{
        position: number;
        sourceTag: string;
        pageNumber?: string;
        prefix?: string;
        suffix?: string;
    }>;
    outputPath?: string;
}
export interface WordIndexArgs {
    filename: string;
    entries: Array<{
        text: string;
        mainEntry: string;
        subEntry?: string;
        pageNumber?: number;
    }>;
    title?: string;
    columns?: number;
    insertAt?: 'end' | 'newPage';
    outputPath?: string;
}
export interface WordMarkIndexEntryArgs {
    filename: string;
    textToMark: string;
    entryText?: string;
    crossReference?: string;
    outputPath?: string;
}
export interface WordFormFieldArgs {
    filename: string;
    fields: Array<{
        type: 'text' | 'checkbox' | 'dropdown' | 'date' | 'number';
        name: string;
        label?: string;
        defaultValue?: string | boolean;
        required?: boolean;
        maxLength?: number;
        helpText?: string;
        options?: string[];
        position?: number;
    }>;
    protectForm?: boolean;
    outputPath?: string;
}
export interface WordContentControlArgs {
    filename: string;
    controls: Array<{
        type: 'richText' | 'plainText' | 'picture' | 'dropDownList' | 'comboBox' | 'datePicker' | 'checkbox';
        title: string;
        tag?: string;
        placeholder?: string;
        options?: string[];
        dateFormat?: string;
        position?: number;
    }>;
    outputPath?: string;
}
export interface WordSmartArtArgs {
    filename: string;
    smartArt: {
        type: 'list' | 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid' | 'picture';
        layout: string;
        items: Array<{
            text: string;
            level?: number;
            image?: string;
        }>;
        colorScheme?: 'colorful' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';
        style?: '3D' | 'flat' | 'cartoon' | 'modern';
    };
    position?: number;
    outputPath?: string;
}
export interface WordEquationArgs {
    filename: string;
    equations: Array<{
        latex?: string;
        mathml?: string;
        text?: string;
        position?: number;
        inline?: boolean;
    }>;
    outputPath?: string;
}
export interface WordSymbolArgs {
    filename: string;
    symbols: Array<{
        character: string;
        font?: string;
        position?: number;
    }>;
    outputPath?: string;
}
export interface WordAccessibilityArgs {
    filename: string;
    checks: {
        altText?: boolean;
        headingStructure?: boolean;
        colorContrast?: boolean;
        tableHeaders?: boolean;
        readingOrder?: boolean;
    };
    autoFix?: boolean;
    outputPath?: string;
}
export interface WordAltTextArgs {
    filename: string;
    altTexts: Array<{
        imageIndex: number;
        altText: string;
        title?: string;
    }>;
    outputPath?: string;
}
export interface WordDigitalSignatureArgs {
    filename: string;
    action: 'add' | 'remove' | 'verify';
    certificatePath?: string;
    password?: string;
    reason?: string;
    location?: string;
    outputPath?: string;
}
export interface WordProtectDocumentArgs {
    filename: string;
    protection: {
        type: 'readOnly' | 'comments' | 'forms' | 'trackedChanges';
        password?: string;
        allowedEditing?: string[];
        users?: string[];
    };
    outputPath?: string;
}
export interface WordMasterDocumentArgs {
    filename: string;
    subdocuments: Array<{
        path: string;
        title?: string;
        lockForEditing?: boolean;
    }>;
    generateTOC?: boolean;
    outputPath?: string;
}
export interface WordAuthorInfoArgs {
    filename: string;
    info: {
        author?: string;
        title?: string;
        subject?: string;
        keywords?: string[];
        category?: string;
        comments?: string;
        company?: string;
        manager?: string;
    };
    outputPath?: string;
}
export interface WordCaptionArgs {
    filename: string;
    captions: Array<{
        type: 'figure' | 'table' | 'equation' | 'custom';
        text: string;
        label?: string;
        numberingFormat?: '1, 2, 3' | 'I, II, III' | 'a, b, c';
        includeChapterNumber?: boolean;
        position?: 'above' | 'below';
        imageOrTableIndex?: number;
    }>;
    outputPath?: string;
}
export interface WordHyperlinkAdvancedArgs {
    filename: string;
    hyperlinks: Array<{
        text: string;
        url?: string;
        bookmark?: string;
        emailAddress?: string;
        screenTip?: string;
        position?: number;
    }>;
    outputPath?: string;
}
export interface WordDropCapArgs {
    filename: string;
    paragraphIndex: number;
    style: 'dropped' | 'inMargin';
    lines?: number;
    distance?: number;
    outputPath?: string;
}
export interface WordWatermarkArgs {
    filename: string;
    watermark: {
        type: 'text' | 'image';
        text?: string;
        imagePath?: string;
        diagonal?: boolean;
        opacity?: number;
        color?: string;
        fontSize?: number;
    };
    outputPath?: string;
}
/**
 * Insert SmartArt graphics with various layouts
 */
export interface PPTSmartArtArgs {
    filename: string;
    slideNumber: number;
    smartArt: {
        type: 'list' | 'process' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid';
        layout: string;
        position: {
            x: number | string;
            y: number | string;
        };
        size?: {
            width: number | string;
            height: number | string;
        };
        items: Array<{
            text: string;
            level?: number;
            children?: string[];
        }>;
        colorScheme?: 'colorful' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';
        style?: '3D' | 'flat' | 'cartoon' | 'outline' | 'filled';
    };
    outputPath?: string;
}
/**
 * Insert SVG icons from Microsoft's icon library
 */
export interface PPTInsertIconsArgs {
    filename: string;
    slideNumber: number;
    icons: Array<{
        name: string;
        category?: 'business' | 'technology' | 'education' | 'analytics' | 'communication' | 'office';
        position: {
            x: number | string;
            y: number | string;
        };
        size?: {
            width: number | string;
            height: number | string;
        };
        color?: string;
        rotation?: number;
    }>;
    outputPath?: string;
}
/**
 * Insert 3D models into presentations
 */
export interface PPTInsert3DModelArgs {
    filename: string;
    slideNumber: number;
    models: Array<{
        path: string;
        position: {
            x: number | string;
            y: number | string;
        };
        size?: {
            width: number | string;
            height: number | string;
        };
        rotation?: {
            x: number;
            y: number;
            z: number;
        };
        animation?: {
            type: 'turntable' | 'swing' | 'jump' | 'none';
            loop?: boolean;
            duration?: number;
        };
        altText?: string;
    }>;
    outputPath?: string;
}
/**
 * Create Zoom links - Summary Zoom, Slide Zoom, or Section Zoom
 */
export interface PPTZoomArgs {
    filename: string;
    slideNumber: number;
    zooms: Array<{
        type: 'summary' | 'slide' | 'section';
        targetSlides?: number[];
        targetSlide?: number;
        targetSection?: string;
        position: {
            x: number | string;
            y: number | string;
        };
        size?: {
            width: number | string;
            height: number | string;
        };
        showReturnToZoom?: boolean;
        useBackground?: boolean;
    }>;
    outputPath?: string;
}
/**
 * Configure screen recording and audio narration settings
 */
export interface PPTRecordingArgs {
    filename: string;
    recording: {
        type: 'screen' | 'audio' | 'both';
        slides?: number[];
        includeNarration?: boolean;
        includeTimings?: boolean;
        includeInkAnnotations?: boolean;
        quality?: 'low' | 'medium' | 'high' | 'HD';
        screenArea?: 'fullScreen' | 'window' | 'custom';
        customArea?: {
            x: number;
            y: number;
            width: number;
            height: number;
        };
    };
    outputPath?: string;
}
/**
 * Embed live web page in a slide
 */
export interface PPTLiveWebArgs {
    filename: string;
    slideNumber: number;
    webPages: Array<{
        url: string;
        position: {
            x: number | string;
            y: number | string;
        };
        size: {
            width: number | string;
            height: number | string;
        };
        refreshInterval?: number;
        allowInteraction?: boolean;
    }>;
    outputPath?: string;
}
/**
 * Apply PowerPoint Designer suggestions (metadata only - actual AI suggestions require PowerPoint)
 */
export interface PPTDesignerArgs {
    filename: string;
    slideNumber?: number;
    preferences?: {
        style?: 'modern' | 'classic' | 'dynamic' | 'minimal';
        layout?: 'auto' | 'imageLeft' | 'imageRight' | 'imageBackground' | 'textOnly';
        colorPalette?: string[];
    };
    outputPath?: string;
}
/**
 * Add comments and @mentions for collaboration
 */
export interface PPTCollaborationArgs {
    filename: string;
    comments: Array<{
        slideNumber: number;
        text: string;
        author?: string;
        position?: {
            x: number;
            y: number;
        };
        mentions?: string[];
        resolved?: boolean;
        replies?: Array<{
            text: string;
            author?: string;
        }>;
    }>;
    outputPath?: string;
}
/**
 * Configure Presenter Coach rehearsal settings (metadata only)
 */
export interface PPTPresenterCoachArgs {
    filename: string;
    settings: {
        enableFeedback?: boolean;
        checkPacing?: boolean;
        checkFillerWords?: boolean;
        checkProfanity?: boolean;
        checkCulturalSensitivity?: boolean;
        checkOriginalPhrases?: boolean;
        checkReadingFromSlide?: boolean;
        targetPace?: number;
    };
    outputPath?: string;
}
/**
 * Configure live subtitles and captions
 */
export interface PPTSubtitlesArgs {
    filename: string;
    subtitles: {
        enable: boolean;
        language?: string;
        position?: 'top' | 'bottom' | 'overlay';
        fontSize?: 'small' | 'medium' | 'large';
        backgroundColor?: string;
        textColor?: string;
        spokenLanguage?: string;
        translationLanguages?: string[];
        showTimestamps?: boolean;
    };
    outputPath?: string;
}
/**
 * Add digital ink drawings and annotations
 */
export interface PPTInkAnnotationsArgs {
    filename: string;
    slideNumber: number;
    annotations: Array<{
        type: 'pen' | 'highlighter' | 'eraser';
        color?: string;
        thickness?: number;
        points: Array<{
            x: number;
            y: number;
        }>;
        pressure?: number[];
    }>;
    outputPath?: string;
}
/**
 * Configure alignment grids and guides for precise object positioning
 */
export interface PPTGridGuidesArgs {
    filename: string;
    slideNumber?: number;
    grid?: {
        show: boolean;
        snapToGrid?: boolean;
        spacing?: number;
    };
    guides?: {
        vertical?: number[];
        horizontal?: number[];
        showGuides?: boolean;
        snapToGuides?: boolean;
    };
    smartGuides?: boolean;
    outputPath?: string;
}
/**
 * Create custom slide shows for different audiences
 */
export interface PPTCustomShowArgs {
    filename: string;
    shows: Array<{
        name: string;
        slides: number[];
        description?: string;
    }>;
    outputPath?: string;
}
/**
 * Manage animation timing and order via Animation Pane
 */
export interface PPTAnimationPaneArgs {
    filename: string;
    slideNumber: number;
    animations: Array<{
        objectId?: string;
        effect: string;
        trigger?: 'onClick' | 'withPrevious' | 'afterPrevious' | 'onPageClick';
        duration?: number;
        delay?: number;
        order?: number;
        repeat?: number | 'untilNextClick' | 'untilEndOfSlide';
        rewind?: boolean;
    }>;
    outputPath?: string;
}
/**
 * Advanced Slide Master customization - fonts, colors, effects, placeholders
 */
export interface PPTSlideMasterAdvancedArgs {
    filename: string;
    master: {
        name?: string;
        theme?: {
            fonts?: {
                heading?: string;
                body?: string;
            };
            colors?: {
                background1?: string;
                background2?: string;
                text1?: string;
                text2?: string;
                accent1?: string;
                accent2?: string;
                accent3?: string;
                accent4?: string;
                accent5?: string;
                accent6?: string;
                hyperlink?: string;
                followedHyperlink?: string;
            };
            effects?: {
                shadow?: boolean;
                reflection?: boolean;
                glow?: boolean;
                softEdges?: boolean;
                threeDFormat?: boolean;
            };
        };
        placeholders?: Array<{
            type: 'title' | 'body' | 'footer' | 'slideNumber' | 'date' | 'subtitle' | 'picture' | 'chart' | 'table' | 'media';
            position: {
                x: number | string;
                y: number | string;
            };
            size: {
                width: number | string;
                height: number | string;
            };
            defaultText?: string;
            formatting?: {
                fontName?: string;
                fontSize?: number;
                bold?: boolean;
                italic?: boolean;
                color?: string;
                alignment?: 'left' | 'center' | 'right' | 'justify';
            };
        }>;
        background?: {
            type: 'solid' | 'gradient' | 'image' | 'pattern';
            color?: string;
            gradient?: {
                colors: string[];
                angle?: number;
            };
            imagePath?: string;
        };
    };
    outputPath?: string;
}
/**
 * Apply or customize presentation themes
 */
export interface PPTThemeArgs {
    filename: string;
    theme: {
        name?: string;
        customThemePath?: string;
        applyToSlides?: number[];
        variants?: 'variant1' | 'variant2' | 'variant3' | 'variant4';
        customizeColors?: {
            accent1?: string;
            accent2?: string;
            accent3?: string;
            accent4?: string;
            accent5?: string;
            accent6?: string;
        };
        customizeFonts?: {
            heading?: string;
            body?: string;
        };
    };
    outputPath?: string;
}
/**
 * Save presentation as a template with placeholders
 */
export interface PPTTemplateArgs {
    filename: string;
    template: {
        title: string;
        description?: string;
        category?: 'business' | 'education' | 'marketing' | 'report' | 'creative' | 'custom';
        placeholders?: Array<{
            slideNumber: number;
            type: 'text' | 'image' | 'chart' | 'table' | 'video';
            label: string;
            position: {
                x: number | string;
                y: number | string;
            };
            size?: {
                width: number | string;
                height: number | string;
            };
            instructions?: string;
        }>;
        protectedElements?: Array<{
            slideNumber: number;
            objectId: string;
        }>;
    };
    outputPath?: string;
}
/**
 * Full IMAP email read with attachments, headers, and metadata
 */
export interface OutlookReadFullEmailArgs {
    messageIds: string[];
    folder?: string;
    includeAttachments?: boolean;
    includeHeaders?: boolean;
    includeRawContent?: boolean;
    markAsRead?: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
    outputPath?: string;
}
/**
 * Delete emails via IMAP
 */
export interface OutlookDeleteEmailArgs {
    messageIds: string[];
    folder?: string;
    permanent?: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Move emails between IMAP folders
 */
export interface OutlookMoveEmailArgs {
    messageIds: string[];
    fromFolder?: string;
    toFolder: string;
    createFolder?: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Create IMAP folder structure
 */
export interface OutlookCreateFolderArgs {
    folderPath: string;
    parent?: string;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Access shared mailboxes and delegate accounts
 */
export interface OutlookSharedMailboxArgs {
    sharedMailbox: string;
    operation: 'list' | 'send' | 'read' | 'move' | 'delete';
    folder?: string;
    messageIds?: string[];
    emailData?: OutlookSendEmailArgs;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
        sharedNamespace?: string;
    };
    outputPath?: string;
}
/**
 * Grant delegate permissions to other users
 */
export interface OutlookDelegateAccessArgs {
    delegateEmail: string;
    permissions: {
        calendar?: 'none' | 'reviewer' | 'author' | 'editor';
        tasks?: 'none' | 'reviewer' | 'author' | 'editor';
        inbox?: 'none' | 'reviewer' | 'author' | 'editor';
        contacts?: 'none' | 'reviewer' | 'author' | 'editor';
        notes?: 'none' | 'reviewer' | 'author' | 'editor';
        journal?: 'none' | 'reviewer' | 'author' | 'editor';
    };
    receiveNotifications?: boolean;
    privateItemsAccess?: boolean;
    outputPath?: string;
}
/**
 * Set automatic replies / out of office / vacation responder
 */
export interface OutlookOutOfOfficeArgs {
    enable: boolean;
    startTime?: string;
    endTime?: string;
    message: {
        internal: string;
        external?: string;
        html?: boolean;
    };
    externalAudience?: 'none' | 'contacts' | 'all';
    declineNewMeetings?: boolean;
    declineMessage?: string;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
    outputPath?: string;
}
/**
 * Create and edit Outlook notes
 */
export interface OutlookNotesArgs {
    notes: Array<{
        subject?: string;
        body: string;
        color?: 'yellow' | 'blue' | 'green' | 'pink' | 'white';
        category?: string;
        createdTime?: string;
    }>;
    outputPath?: string;
}
/**
 * Create journal entries for activity tracking
 */
export interface OutlookJournalArgs {
    entries: Array<{
        subject: string;
        entryType: 'phoneCall' | 'email' | 'meeting' | 'task' | 'document' | 'note' | 'custom';
        startTime: string;
        duration?: number;
        description?: string;
        contacts?: string[];
        categories?: string[];
        company?: string;
    }>;
    outputPath?: string;
}
/**
 * Subscribe to and manage RSS feeds
 */
export interface OutlookRSSFeedArgs {
    operation: 'add' | 'remove' | 'list' | 'update';
    feeds?: Array<{
        url: string;
        name?: string;
        folder?: string;
        updateInterval?: number;
        maxItems?: number;
    }>;
    outputPath?: string;
}
/**
 * Manage PST/OST data files (metadata operations only)
 */
export interface OutlookDataFileArgs {
    operation: 'create' | 'open' | 'close' | 'compact' | 'info';
    filePath?: string;
    fileType?: 'pst' | 'ost';
    password?: string;
    displayName?: string;
    deliverToThisFile?: boolean;
    outputPath?: string;
}
/**
 * Create Quick Steps for email automation
 */
export interface OutlookQuickStepsArgs {
    steps: Array<{
        name: string;
        description?: string;
        actions: Array<{
            type: 'move' | 'categorize' | 'flag' | 'forward' | 'reply' | 'delete' | 'markRead' | 'markUnread';
            folder?: string;
            category?: string;
            flagType?: 'today' | 'tomorrow' | 'thisWeek' | 'nextWeek' | 'noDate' | 'complete';
            forwardTo?: string[];
            replyTemplate?: string;
        }>;
        shortcut?: string;
    }>;
    outputPath?: string;
}
/**
 * Configure conversation view settings
 */
export interface OutlookConversationViewArgs {
    enable: boolean;
    settings?: {
        showMessagesFromOtherFolders?: boolean;
        showSenders?: boolean;
        alwaysExpand?: boolean;
        useClassicIndentation?: boolean;
        highlightUnread?: boolean;
    };
    folders?: string[];
    outputPath?: string;
}
/**
 * Clean up redundant messages in conversations
 */
export interface OutlookCleanupArgs {
    folder?: string;
    scope: 'conversation' | 'folderAndSubfolders' | 'selectedMessages';
    messageIds?: string[];
    deleteRedundant?: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Ignore conversation threads
 */
export interface OutlookIgnoreConversationArgs {
    conversationIds: string[];
    restore?: boolean;
    deleteExisting?: boolean;
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Flag emails with colors and due dates
 */
export interface OutlookFlagEmailArgs {
    messageIds: string[];
    flag: {
        type: 'today' | 'tomorrow' | 'thisWeek' | 'nextWeek' | 'noDate' | 'complete' | 'clear';
        dueDate?: string;
        startDate?: string;
        reminderTime?: string;
        color?: 'red' | 'blue' | 'yellow' | 'green' | 'orange' | 'purple';
    };
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
}
/**
 * Create and apply color categories
 */
export interface OutlookCategoryArgs {
    operation: 'create' | 'apply' | 'remove' | 'list';
    categories?: Array<{
        name: string;
        color: 'red' | 'orange' | 'yellow' | 'green' | 'blue' | 'purple' | 'pink' | 'brown' | 'gray';
        shortcut?: string;
    }>;
    messageIds?: string[];
    categoryNames?: string[];
    imapConfig?: {
        host: string;
        port: number;
        user: string;
        password: string;
        tls?: boolean;
    };
    outputPath?: string;
}
/**
 * Create HTML email signatures with images and formatting
 */
export interface OutlookSignatureArgs {
    signatures: Array<{
        name: string;
        html: string;
        text?: string;
        images?: Array<{
            filename: string;
            path: string;
            cid?: string;
        }>;
        defaultFor?: {
            newMessages?: boolean;
            replies?: boolean;
        };
    }>;
    outputPath?: string;
}
/**
 * Manage autocomplete nickname cache
 */
export interface OutlookAutoCompleteArgs {
    operation: 'list' | 'add' | 'remove' | 'clear' | 'export' | 'import';
    entries?: Array<{
        name: string;
        email: string;
        lastUsed?: string;
    }>;
    filePath?: string;
    outputPath?: string;
}
/**
 * Advanced mail merge with filters and conditional content
 */
export interface OutlookMailMergeAdvancedArgs {
    templatePath?: string;
    template?: {
        subject: string;
        body: string;
        html?: boolean;
    };
    dataSource: Record<string, string | number>[];
    filters?: Array<{
        field: string;
        operator: 'equals' | 'notEquals' | 'contains' | 'greaterThan' | 'lessThan' | 'startsWith' | 'endsWith';
        value: string | number;
    }>;
    conditionalContent?: Array<{
        condition: string;
        content: string;
    }>;
    attachments?: Array<{
        filename: string;
        path?: string;
        conditional?: string;
    }>;
    sendOptions?: {
        batchSize?: number;
        delayBetweenBatches?: number;
        testMode?: boolean;
        testAddress?: string;
    };
    smtpConfig?: OutlookSMTPConfig;
    outputPath?: string;
}
//# sourceMappingURL=types.d.ts.map