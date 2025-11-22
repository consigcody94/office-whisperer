/**
 * Word Generator - Create Word documents using docx library
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableCell,
  TableRow,
  HeadingLevel,
  AlignmentType,
  PageBreak,
  UnderlineType,
  TableOfContents,
  Header,
  Footer,
  ImageRun,
} from 'docx';
import * as fs from 'fs/promises';
import type {
  WordDocumentOptions,
  WordSection,
  WordElement,
  WordParagraph,
  WordTable,
  WordStyles,
} from '../types.js';

export class WordGenerator {
  async createDocument(options: WordDocumentOptions): Promise<Buffer> {
    const sections = options.sections.map(section => ({
      properties: section.properties || {},
      children: this.processElements(section.children),
    }));

    const doc = new Document({
      sections,
      styles: options.styles as any,
    });

    return await Packer.toBuffer(doc);
  }

  async addTableOfContents(
    filename: string,
    title?: string,
    hyperlinks: boolean = true,
    levels: number = 3
  ): Promise<Buffer> {
    const existingBuffer = await this.loadDocument(filename);

    // Create a new document with TOC
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: title || 'Table of Contents',
            heading: HeadingLevel.HEADING_1,
          }),
          new TableOfContents('Table of Contents', {
            hyperlink: hyperlinks,
            headingStyleRange: `1-${levels}`,
          }),
          new Paragraph({ pageBreakBefore: true }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async mailMerge(
    templatePath: string,
    dataSource: Record<string, string | number>[],
    outputFilename?: string
  ): Promise<Buffer[]> {
    // Load template (in production)
    // For now, create merged documents
    const documents: Buffer[] = [];

    for (const data of dataSource) {
      const doc = new Document({
        sections: [{
          children: [
            new Paragraph({
              text: 'Mail Merge Document',
              heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
              text: `Generated from template: ${templatePath}`,
            }),
            new Paragraph({
              text: `Data: ${JSON.stringify(data)}`,
            }),
          ],
        }],
      });

      documents.push(await Packer.toBuffer(doc));
    }

    return documents;
  }

  async findReplace(
    filename: string,
    find: string,
    replace: string,
    matchCase: boolean = false,
    matchWholeWord: boolean = false,
    formatting?: any
  ): Promise<Buffer> {
    // Note: docx library doesn't support loading existing documents for editing
    // This would require using a different library like docx4js or pizzip
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: `Find and Replace: "${find}" -> "${replace}"`,
          }),
          new Paragraph({
            text: `Match Case: ${matchCase}, Match Whole Word: ${matchWholeWord}`,
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async addComment(
    filename: string,
    text: string,
    comment: string,
    author: string = 'Office Whisperer'
  ): Promise<Buffer> {
    // Comments require more advanced features
    // Creating a placeholder document
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: `Text: ${text}`,
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `Comment by ${author}: ${comment}`,
                italics: true,
                color: '0000FF',
              }),
            ],
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async formatStyles(
    filename: string,
    styles: WordStyles
  ): Promise<Buffer> {
    const doc = new Document({
      styles: styles as any,
      sections: [{
        children: [
          new Paragraph({
            text: 'Document with Custom Styles Applied',
            heading: HeadingLevel.HEADING_1,
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async insertImage(
    filename: string,
    imagePath: string,
    position?: { x?: number; y?: number },
    size?: { width?: number; height?: number },
    wrapping?: string
  ): Promise<Buffer> {
    const imageBuffer = await fs.readFile(imagePath);

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: 'Document with Image',
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            children: [
              new ImageRun({
                data: imageBuffer,
                transformation: {
                  width: size?.width || 200,
                  height: size?.height || 200,
                },
              }),
            ],
          }),
          new Paragraph({
            text: `Image inserted from: ${imagePath}`,
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async addHeaderFooter(
    filename: string,
    type: 'header' | 'footer',
    content: WordElement[],
    sectionType: 'default' | 'first' | 'even' = 'default'
  ): Promise<Buffer> {
    const processedContent = this.processElements(content);

    const sectionConfig: any = {
      children: [
        new Paragraph({
          text: 'Document with Custom Header/Footer',
        }),
      ],
    };

    if (type === 'header') {
      sectionConfig.headers = {
        default: new Header({
          children: processedContent,
        }),
      };
    } else {
      sectionConfig.footers = {
        default: new Footer({
          children: processedContent,
        }),
      };
    }

    const doc = new Document({
      sections: [sectionConfig],
    });

    return await Packer.toBuffer(doc);
  }

  async compareDocuments(
    originalPath: string,
    revisedPath: string,
    author: string = 'Office Whisperer'
  ): Promise<Buffer> {
    // Document comparison would require specialized libraries
    // Creating a comparison report
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: 'Document Comparison Report',
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            text: `Original: ${originalPath}`,
          }),
          new Paragraph({
            text: `Revised: ${revisedPath}`,
          }),
          new Paragraph({
            text: `Reviewer: ${author}`,
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Comparison analysis would be performed here...',
                italics: true,
              }),
            ],
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async convertToPDF(filename: string): Promise<Buffer> {
    // PDF conversion would require LibreOffice or similar
    // Return placeholder info
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: 'PDF Conversion Information',
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            text: `Source document: ${filename}`,
          }),
          new Paragraph({
            text: 'PDF conversion requires external tools like LibreOffice or docx2pdf',
          }),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  async mergeDocuments(documentPaths: string[], outputPath: string): Promise<Buffer> {
    // In production, this would load and merge multiple documents
    // For now, create a placeholder document
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: 'Merged Documents',
            heading: HeadingLevel.HEADING_1,
          }),
          ...documentPaths.map(path => new Paragraph({ text: `- ${path}` })),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }

  // Helper methods
  private processElements(elements: WordElement[]): any[] {
    const processed: any[] = [];

    for (const element of elements) {
      switch (element.type) {
        case 'paragraph':
          processed.push(this.createParagraph(element));
          break;
        case 'table':
          processed.push(this.createTable(element));
          break;
        case 'pageBreak':
          processed.push(new Paragraph({ pageBreakBefore: true }));
          break;
        case 'image':
          processed.push(new Paragraph({ text: `[Image: ${element.path}]` }));
          break;
        case 'toc':
          processed.push(
            new Paragraph({
              text: element.title || 'Table of Contents',
              heading: HeadingLevel.HEADING_1,
            })
          );
          processed.push(
            new TableOfContents('Table of Contents', {
              hyperlink: true,
              headingStyleRange: '1-3',
            })
          );
          break;
      }
    }

    return processed;
  }

  private createParagraph(para: WordParagraph): Paragraph {
    const config: any = {
      alignment: this.getAlignment(para.alignment),
    };

    // Handle heading levels
    if (para.heading) {
      const level = para.heading.replace('Heading', '');
      config.heading = HeadingLevel[`HEADING_${level}` as keyof typeof HeadingLevel];
    }

    // Handle bullet points
    if (para.bullet) {
      config.bullet = { level: para.bullet.level };
    }

    // Handle numbering
    if (para.numbering) {
      config.numbering = para.numbering;
    }

    // Handle spacing
    if (para.spacing) {
      config.spacing = para.spacing;
    }

    // Create text runs if children specified
    if (para.children && para.children.length > 0) {
      config.children = para.children.map(run => new TextRun({
        text: run.text,
        bold: run.bold,
        italics: run.italics,
        underline: run.underline ? { type: UnderlineType.SINGLE } : undefined,
        strike: run.strike,
        color: run.color,
        size: run.size ? run.size * 2 : undefined, // Half-points in docx
        font: run.font,
        highlight: run.highlight,
      }));
    } else if (para.text) {
      config.text = para.text;
    }

    return new Paragraph(config);
  }

  private createTable(table: WordTable): Table {
    const rows = table.rows.map(row => new TableRow({
      children: row.cells.map(cell => new TableCell({
        children: cell.children.map(p => this.createParagraph(p)),
        shading: cell.shading as any,
        margins: cell.margins as any,
        columnSpan: cell.columnSpan,
        rowSpan: cell.rowSpan,
      })),
      height: row.height ? { value: row.height, rule: 'exact' as const } : undefined,
      cantSplit: row.cantSplit,
      tableHeader: row.tableHeader,
    }));

    return new Table({
      rows,
      width: table.width as any,
      borders: table.borders as any,
    });
  }

  private getAlignment(align?: string): typeof AlignmentType[keyof typeof AlignmentType] {
    switch (align) {
      case 'center':
        return AlignmentType.CENTER;
      case 'right':
        return AlignmentType.RIGHT;
      case 'justified':
        return AlignmentType.JUSTIFIED;
      default:
        return AlignmentType.LEFT;
    }
  }

  private async loadDocument(filename: string): Promise<Buffer> {
    try {
      return await fs.readFile(filename);
    } catch (error) {
      console.warn(`File ${filename} not found, creating new document`);
      return Buffer.from([]);
    }
  }
}
