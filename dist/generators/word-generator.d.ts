/**
 * Word Generator - Create Word documents using docx library
 */
import type { WordDocumentOptions } from '../types.js';
export declare class WordGenerator {
    createDocument(options: WordDocumentOptions): Promise<Buffer>;
    private processElements;
    private createParagraph;
    private createTable;
    private getAlignment;
    mergeDocuments(documentPaths: string[], outputPath: string): Promise<Buffer>;
}
//# sourceMappingURL=word-generator.d.ts.map