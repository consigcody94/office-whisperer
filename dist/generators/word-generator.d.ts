/**
 * Word Generator - Create Word documents using docx library
 */
import type { WordDocumentOptions, WordElement, WordStyles } from '../types.js';
export declare class WordGenerator {
    createDocument(options: WordDocumentOptions): Promise<Buffer>;
    addTableOfContents(filename: string, title?: string, hyperlinks?: boolean, levels?: number): Promise<Buffer>;
    mailMerge(templatePath: string, dataSource: Record<string, string | number>[], outputFilename?: string): Promise<Buffer[]>;
    findReplace(filename: string, find: string, replace: string, matchCase?: boolean, matchWholeWord?: boolean, formatting?: any): Promise<Buffer>;
    addComment(filename: string, text: string, comment: string, author?: string): Promise<Buffer>;
    formatStyles(filename: string, styles: WordStyles): Promise<Buffer>;
    insertImage(filename: string, imagePath: string, position?: {
        x?: number;
        y?: number;
    }, size?: {
        width?: number;
        height?: number;
    }, wrapping?: string): Promise<Buffer>;
    addHeaderFooter(filename: string, type: 'header' | 'footer', content: WordElement[], sectionType?: 'default' | 'first' | 'even'): Promise<Buffer>;
    compareDocuments(originalPath: string, revisedPath: string, author?: string): Promise<Buffer>;
    convertToPDF(filename: string): Promise<Buffer>;
    mergeDocuments(documentPaths: string[], outputPath: string): Promise<Buffer>;
    private processElements;
    private createParagraph;
    private createTable;
    private getAlignment;
    private loadDocument;
    enableTrackChanges(filename: string, enable: boolean, author?: string): Promise<Buffer>;
    addFootnotes(filename: string, footnotes: Array<{
        text: string;
        note: string;
        type?: 'footnote' | 'endnote';
    }>): Promise<Buffer>;
    addBookmarks(filename: string, bookmarks: Array<{
        name: string;
        text: string;
    }>): Promise<Buffer>;
    addSectionBreaks(filename: string, breaks: Array<{
        position: number;
        type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
    }>): Promise<Buffer>;
    addTextBoxes(filename: string, textBoxes: Array<{
        text: string;
        position?: {
            x: number;
            y: number;
        };
        width?: number;
        height?: number;
    }>): Promise<Buffer>;
    addCrossReferences(filename: string, references: Array<{
        bookmarkName: string;
        referenceType: 'pageNumber' | 'text' | 'above/below';
        insertText?: string;
    }>): Promise<Buffer>;
    /**
     * Create bibliography from sources with citation styles
     * Note: Uses simplified bibliography format as docx library doesn't support native Word bibliography
     */
    addBibliography(filename: string, sources: Array<{
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
    }>, style?: 'APA' | 'MLA' | 'Chicago' | 'Harvard' | 'IEEE', insertAt?: 'end' | 'cursor'): Promise<Buffer>;
    /**
     * Insert citations referencing bibliography sources
     * Note: Creates inline citations as text; full citation features require Microsoft Word
     */
    insertCitations(filename: string, citations: Array<{
        position: number;
        sourceTag: string;
        pageNumber?: string;
        prefix?: string;
        suffix?: string;
    }>): Promise<Buffer>;
    /**
     * Generate index from entries with columns and formatting
     * Note: Creates a text-based index; automatic index generation requires Microsoft Word
     */
    createIndex(filename: string, entries: Array<{
        text: string;
        mainEntry: string;
        subEntry?: string;
        pageNumber?: number;
    }>, title?: string, columns?: number, insertAt?: 'end' | 'newPage'): Promise<Buffer>;
    /**
     * Mark text for automatic index generation
     * Note: Creates marked text notation; automatic indexing requires Microsoft Word
     */
    markIndexEntry(filename: string, textToMark: string, entryText?: string, crossReference?: string): Promise<Buffer>;
    /**
     * Add form fields with protection
     * Note: Creates form field representations; full form functionality requires Microsoft Word
     */
    addFormFields(filename: string, fields: Array<{
        type: 'text' | 'checkbox' | 'dropdown' | 'date' | 'number';
        name: string;
        label?: string;
        defaultValue?: string | boolean;
        required?: boolean;
        maxLength?: number;
        helpText?: string;
        options?: string[];
        position?: number;
    }>, protectForm?: boolean): Promise<Buffer>;
    /**
     * Add content controls (rich text, picture, date picker)
     * Note: Creates content control placeholders; full controls require Microsoft Word
     */
    addContentControls(filename: string, controls: Array<{
        type: 'richText' | 'plainText' | 'picture' | 'dropDownList' | 'comboBox' | 'datePicker' | 'checkbox';
        title: string;
        tag?: string;
        placeholder?: string;
        options?: string[];
        dateFormat?: string;
        position?: number;
    }>): Promise<Buffer>;
    /**
     * Insert SmartArt graphics as placeholders
     * Note: SmartArt rendering not supported by docx library; creates descriptive placeholders
     */
    insertSmartArt(filename: string, smartArt: {
        type: 'list' | 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid' | 'picture';
        layout: string;
        items: Array<{
            text: string;
            level?: number;
            image?: string;
        }>;
        colorScheme?: 'colorful' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';
        style?: '3D' | 'flat' | 'cartoon' | 'modern';
    }, position?: number): Promise<Buffer>;
    /**
     * Insert mathematical equations as placeholders
     * Note: Equation rendering not fully supported; creates text representations
     */
    addEquations(filename: string, equations: Array<{
        latex?: string;
        mathml?: string;
        text?: string;
        position?: number;
        inline?: boolean;
    }>): Promise<Buffer>;
    /**
     * Insert special characters and symbols
     */
    insertSymbols(filename: string, symbols: Array<{
        character: string;
        font?: string;
        position?: number;
    }>): Promise<Buffer>;
    /**
     * Check document for accessibility issues
     * Note: Limited accessibility checking; full analysis requires Microsoft Word
     */
    checkAccessibility(filename: string, checks: {
        altText?: boolean;
        headingStructure?: boolean;
        colorContrast?: boolean;
        tableHeaders?: boolean;
        readingOrder?: boolean;
    }, autoFix?: boolean): Promise<Buffer>;
    /**
     * Add alt text to images for accessibility
     * Note: Creates alt text metadata; requires proper image handling
     */
    setAltText(filename: string, altTexts: Array<{
        imageIndex: number;
        altText: string;
        title?: string;
    }>): Promise<Buffer>;
    /**
     * Add digital signature metadata
     * Note: Limited to metadata only; cryptographic signing requires external tools
     */
    addDigitalSignature(filename: string, action: 'add' | 'remove' | 'verify', certificatePath?: string, password?: string, reason?: string, location?: string): Promise<Buffer>;
    /**
     * Set document protection
     * Note: Creates protection metadata; full enforcement requires Microsoft Word
     */
    protectDocument(filename: string, protection: {
        type: 'readOnly' | 'comments' | 'forms' | 'trackedChanges';
        password?: string;
        allowedEditing?: string[];
        users?: string[];
    }): Promise<Buffer>;
    /**
     * Create master document with subdocuments
     * Note: Creates master document structure; subdocument linking requires Microsoft Word
     */
    createMasterDocument(filename: string, subdocuments: Array<{
        path: string;
        title?: string;
        lockForEditing?: boolean;
    }>, generateTOC?: boolean): Promise<Buffer>;
    /**
     * Set document metadata (author, title, subject, keywords, company)
     */
    setDocumentInfo(filename: string, info: {
        author?: string;
        title?: string;
        subject?: string;
        keywords?: string[];
        category?: string;
        comments?: string;
        company?: string;
        manager?: string;
    }): Promise<Buffer>;
    /**
     * Add captions to figures, tables, or equations
     * Note: Creates caption text; automatic caption numbering requires Microsoft Word
     */
    addCaptions(filename: string, captions: Array<{
        type: 'figure' | 'table' | 'equation' | 'custom';
        text: string;
        label?: string;
        numberingFormat?: '1, 2, 3' | 'I, II, III' | 'a, b, c';
        includeChapterNumber?: boolean;
        position?: 'above' | 'below';
        imageOrTableIndex?: number;
    }>): Promise<Buffer>;
    /**
     * Add advanced hyperlinks with bookmarks, mailto, and tooltips
     */
    addAdvancedHyperlinks(filename: string, hyperlinks: Array<{
        text: string;
        url?: string;
        bookmark?: string;
        emailAddress?: string;
        screenTip?: string;
        position?: number;
    }>): Promise<Buffer>;
    /**
     * Add drop cap to paragraph
     * Note: Drop cap formatting not fully supported; creates styled first letter
     */
    addDropCap(filename: string, paragraphIndex: number, style: 'dropped' | 'inMargin', lines?: number, distance?: number): Promise<Buffer>;
    /**
     * Add text or image watermark
     * Note: Watermark not fully supported by docx library; creates placeholder
     */
    addWatermark(filename: string, watermark: {
        type: 'text' | 'image';
        text?: string;
        imagePath?: string;
        diagonal?: boolean;
        opacity?: number;
        color?: string;
        fontSize?: number;
    }): Promise<Buffer>;
    private formatAPA;
    private formatMLA;
    private formatChicago;
    private formatHarvard;
    private formatIEEE;
    private formatCaptionNumber;
    private toRoman;
}
//# sourceMappingURL=word-generator.d.ts.map