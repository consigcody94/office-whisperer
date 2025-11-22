/**
 * Word Generator - Create Word documents using docx library
 */
import { Document, Packer, Paragraph, TextRun, Table, TableCell, TableRow, HeadingLevel, AlignmentType, UnderlineType, TableOfContents, Header, Footer, ImageRun, } from 'docx';
import * as fs from 'fs/promises';
export class WordGenerator {
    async createDocument(options) {
        const sections = options.sections.map(section => ({
            properties: section.properties || {},
            children: this.processElements(section.children),
        }));
        const doc = new Document({
            sections,
            styles: options.styles,
        });
        return await Packer.toBuffer(doc);
    }
    async addTableOfContents(filename, title, hyperlinks = true, levels = 3) {
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
    async mailMerge(templatePath, dataSource, outputFilename) {
        // Load template (in production)
        // For now, create merged documents
        const documents = [];
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
    async findReplace(filename, find, replace, matchCase = false, matchWholeWord = false, formatting) {
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
    async addComment(filename, text, comment, author = 'Office Whisperer') {
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
    async formatStyles(filename, styles) {
        const doc = new Document({
            styles: styles,
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
    async insertImage(filename, imagePath, position, size, wrapping) {
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
    async addHeaderFooter(filename, type, content, sectionType = 'default') {
        const processedContent = this.processElements(content);
        const sectionConfig = {
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
        }
        else {
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
    async compareDocuments(originalPath, revisedPath, author = 'Office Whisperer') {
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
    async convertToPDF(filename) {
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
    async mergeDocuments(documentPaths, outputPath) {
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
    processElements(elements) {
        const processed = [];
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
                    processed.push(new Paragraph({
                        text: element.title || 'Table of Contents',
                        heading: HeadingLevel.HEADING_1,
                    }));
                    processed.push(new TableOfContents('Table of Contents', {
                        hyperlink: true,
                        headingStyleRange: '1-3',
                    }));
                    break;
            }
        }
        return processed;
    }
    createParagraph(para) {
        const config = {
            alignment: this.getAlignment(para.alignment),
        };
        // Handle heading levels
        if (para.heading) {
            const level = para.heading.replace('Heading', '');
            config.heading = HeadingLevel[`HEADING_${level}`];
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
        }
        else if (para.text) {
            config.text = para.text;
        }
        return new Paragraph(config);
    }
    createTable(table) {
        const rows = table.rows.map(row => new TableRow({
            children: row.cells.map(cell => new TableCell({
                children: cell.children.map(p => this.createParagraph(p)),
                shading: cell.shading,
                margins: cell.margins,
                columnSpan: cell.columnSpan,
                rowSpan: cell.rowSpan,
            })),
            height: row.height ? { value: row.height, rule: 'exact' } : undefined,
            cantSplit: row.cantSplit,
            tableHeader: row.tableHeader,
        }));
        return new Table({
            rows,
            width: table.width,
            borders: table.borders,
        });
    }
    getAlignment(align) {
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
    async loadDocument(filename) {
        try {
            return await fs.readFile(filename);
        }
        catch (error) {
            console.warn(`File ${filename} not found, creating new document`);
            return Buffer.from([]);
        }
    }
    // ============================================================================
    // Word v3.0 Methods - Phase 1 Quick Wins
    // ============================================================================
    async enableTrackChanges(filename, enable, author) {
        // Note: docx library has limited track changes support
        // Creating a document with metadata about track changes settings
        const doc = new Document({
            sections: [{
                    children: [
                        new Paragraph({
                            text: `Track Changes: ${enable ? 'ENABLED' : 'DISABLED'}`,
                            heading: HeadingLevel.HEADING_1,
                        }),
                        new Paragraph({
                            text: `Author: ${author || 'Office Whisperer'}`,
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Note: Full track changes functionality requires Microsoft Word.',
                                    italics: true,
                                    color: '666666',
                                }),
                            ],
                        }),
                    ],
                }],
        });
        return await Packer.toBuffer(doc);
    }
    async addFootnotes(filename, footnotes) {
        const children = [];
        footnotes.forEach((fn, index) => {
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: fn.text }),
                    new TextRun({
                        text: ` [${index + 1}]`,
                        superScript: true,
                        color: '0000FF',
                    }),
                ],
            }));
        });
        // Add footnotes section
        children.push(new Paragraph({ pageBreakBefore: true }));
        children.push(new Paragraph({
            text: 'Footnotes',
            heading: HeadingLevel.HEADING_2,
        }));
        footnotes.forEach((fn, index) => {
            children.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `${index + 1}. `,
                        superScript: true,
                    }),
                    new TextRun({
                        text: fn.note,
                        size: 20,
                    }),
                ],
            }));
        });
        const doc = new Document({
            sections: [{ children }],
        });
        return await Packer.toBuffer(doc);
    }
    async addBookmarks(filename, bookmarks) {
        // Note: docx library has limited bookmark support
        const children = [
            new Paragraph({
                text: 'Document with Bookmarks',
                heading: HeadingLevel.HEADING_1,
            }),
        ];
        bookmarks.forEach(bookmark => {
            children.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `📑 Bookmark "${bookmark.name}": `,
                        bold: true,
                    }),
                    new TextRun({
                        text: bookmark.text,
                    }),
                ],
            }));
        });
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: '\nNote: Full bookmark functionality requires Microsoft Word.',
                    italics: true,
                    color: '666666',
                }),
            ],
        }));
        const doc = new Document({
            sections: [{ children }],
        });
        return await Packer.toBuffer(doc);
    }
    async addSectionBreaks(filename, breaks) {
        const sections = [];
        breaks.forEach((brk, index) => {
            sections.push({
                properties: {
                    type: brk.type === 'nextPage' ? 'nextPage' : 'continuous',
                },
                children: [
                    new Paragraph({
                        text: `Section ${index + 1}`,
                        heading: HeadingLevel.HEADING_1,
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `Break Type: ${brk.type}`,
                                italics: true,
                            }),
                        ],
                    }),
                ],
            });
        });
        const doc = new Document({ sections });
        return await Packer.toBuffer(doc);
    }
    async addTextBoxes(filename, textBoxes) {
        // Note: docx library has limited text box support
        const children = [
            new Paragraph({
                text: 'Document with Text Boxes',
                heading: HeadingLevel.HEADING_1,
            }),
        ];
        textBoxes.forEach((box, index) => {
            children.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `\n[Text Box ${index + 1}]`,
                        bold: true,
                        color: '0066CC',
                    }),
                ],
            }), new Paragraph({
                text: box.text,
                border: {
                    top: { style: 'single', size: 1, color: '0066CC' },
                    bottom: { style: 'single', size: 1, color: '0066CC' },
                    left: { style: 'single', size: 1, color: '0066CC' },
                    right: { style: 'single', size: 1, color: '0066CC' },
                },
                shading: {
                    fill: 'F0F8FF',
                },
            }));
        });
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: '\nNote: Full text box positioning requires Microsoft Word.',
                    italics: true,
                    color: '666666',
                }),
            ],
        }));
        const doc = new Document({
            sections: [{ children }],
        });
        return await Packer.toBuffer(doc);
    }
    async addCrossReferences(filename, references) {
        const children = [
            new Paragraph({
                text: 'Document with Cross-References',
                heading: HeadingLevel.HEADING_1,
            }),
        ];
        references.forEach(ref => {
            children.push(new Paragraph({
                children: [
                    new TextRun({
                        text: ref.insertText || 'See ',
                    }),
                    new TextRun({
                        text: `"${ref.bookmarkName}"`,
                        bold: true,
                        color: '0000FF',
                    }),
                    new TextRun({
                        text: ` (${ref.referenceType})`,
                        italics: true,
                    }),
                ],
            }));
        });
        children.push(new Paragraph({
            children: [
                new TextRun({
                    text: '\nNote: Full cross-reference linking requires Microsoft Word.',
                    italics: true,
                    color: '666666',
                }),
            ],
        }));
        const doc = new Document({
            sections: [{ children }],
        });
        return await Packer.toBuffer(doc);
    }
}
//# sourceMappingURL=word-generator.js.map