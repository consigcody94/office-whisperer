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
} from 'docx';
import type { WordDocumentOptions, WordSection, WordElement, WordParagraph, WordTable } from '../types.js';

export class WordGenerator {
  async createDocument(options: WordDocumentOptions): Promise<Buffer> {
    const sections = options.sections.map(section => ({
      properties: section.properties || {},
      children: this.processElements(section.children),
    }));

    const doc = new Document({
      sections,
    });

    return await Packer.toBuffer(doc);
  }

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
          processed.push(new PageBreak());
          break;
        case 'image':
          // Image handling would require fs access
          processed.push(new Paragraph({ text: `[Image: ${element.path}]` }));
          break;
        case 'toc':
          processed.push(new Paragraph({ text: '[Table of Contents]' }));
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

  async mergeDocuments(documentPaths: string[], outputPath: string): Promise<Buffer> {
    // In production, this would load and merge multiple documents
    // For now, create a placeholder document
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({ text: 'Merged Documents:' }),
          ...documentPaths.map(path => new Paragraph({ text: `- ${path}` })),
        ],
      }],
    });

    return await Packer.toBuffer(doc);
  }
}
