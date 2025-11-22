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
}
//# sourceMappingURL=word-generator.d.ts.map