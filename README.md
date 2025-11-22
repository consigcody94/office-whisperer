# 🎯 Office Whisperer

**Control Microsoft Office Suite through natural language with Claude Desktop**

Transform your Office workflow with AI-powered automation. Create Excel spreadsheets, Word documents, and PowerPoint presentations using simple conversation.

[![TypeScript](https://img.shields.io/badge/TypeScript-5.3-blue?logo=typescript)](https://www.typescriptlang.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-2025--06--18-purple)](https://modelcontextprotocol.io)

## ✨ Features

### 📊 Excel Automation
- **Smart Workbooks** - Create multi-sheet workbooks with data, formulas, and formatting
- **Advanced Formulas** - SUM, AVERAGE, VLOOKUP, and complex calculations
- **Beautiful Charts** - Line, bar, pie, scatter, and area charts
- **CSV Export** - Convert Excel files to CSV format
- **Pivot Tables** - Dynamic data summarization and analysis
- **Auto-Formatting** - Professional styling with headers, colors, and borders

### 📄 Word Documents
- **Professional Documents** - Multi-section documents with headers and footers
- **Rich Formatting** - Bold, italic, underline, colors, fonts, and sizes
- **Tables & Lists** - Structured data with borders and shading
- **Headings & TOC** - Automatic heading levels and table of contents
- **Page Breaks** - Control document flow and pagination
- **Merge Documents** - Combine multiple documents into one

### 🎬 PowerPoint Presentations
- **Beautiful Slides** - Title, content, section, comparison, and blank layouts
- **Custom Themes** - Light, dark, colorful, or default themes
- **Rich Content** - Text, images, shapes, tables, and charts
- **Speaker Notes** - Add presenter notes to each slide
- **Backgrounds** - Colors and images for slide backgrounds
- **Export Options** - Multiple formats and sizes

## 🚀 Quick Start

### Installation

```bash
git clone https://github.com/consigcody94/office-whisperer.git
cd office-whisperer
npm install
npm run build
```

### Claude Desktop Setup

Add to your `claude_desktop_config.json`:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**Linux:** `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "office-whisperer": {
      "command": "node",
      "args": ["/path/to/office-whisperer/dist/mcp-server.js"]
    }
  }
}
```

Restart Claude Desktop.

## 💬 Usage Examples

### Excel - Sales Report

> "Create a sales report Excel file with monthly revenue data, a totals row with SUM formulas, and a chart showing the trend"

```javascript
{
  "filename": "sales_report.xlsx",
  "sheets": [{
    "name": "Q4 Sales",
    "columns": [
      { "header": "Month", "key": "month", "width": 15 },
      { "header": "Revenue", "key": "revenue", "width": 15 },
      { "header": "Expenses", "key": "expenses", "width": 15 },
      { "header": "Profit", "key": "profit", "width": 15 }
    ],
    "data": [
      ["October", 45000, 12000, 33000],
      ["November", 52000, 14000, 38000],
      ["December", 68000, 16000, 52000]
    ]
  }]
}
```

### Word - Project Proposal

> "Create a project proposal document with a title page, executive summary section, and a table showing project milestones"

```javascript
{
  "filename": "proposal.docx",
  "title": "AI Integration Project Proposal",
  "sections": [{
    "children": [
      {
        "type": "paragraph",
        "text": "AI Integration Project Proposal",
        "heading": "Heading1"
      },
      {
        "type": "paragraph",
        "text": "Transforming customer service with artificial intelligence",
        "alignment": "center"
      },
      {
        "type": "table",
        "rows": [
          {
            "cells": [
              { "children": [{ "type": "paragraph", "text": "Milestone" }] },
              { "children": [{ "type": "paragraph", "text": "Deadline" }] }
            ],
            "tableHeader": true
          },
          {
            "cells": [
              { "children": [{ "type": "paragraph", "text": "Requirements" }] },
              { "children": [{ "type": "paragraph", "text": "Week 1" }] }
            ]
          }
        ]
      }
    ]
  }]
}
```

### PowerPoint - Quarterly Review

> "Make a quarterly review presentation with a title slide, 3 content slides with bullet points, and a thank you slide"

```javascript
{
  "filename": "q4_review.pptx",
  "title": "Q4 2024 Review",
  "theme": "dark",
  "slides": [
    {
      "layout": "title",
      "title": "Q4 2024 Quarterly Review",
      "subtitle": "Strong Growth Across All Metrics"
    },
    {
      "layout": "content",
      "title": "Key Achievements",
      "content": [
        {
          "type": "text",
          "text": "• 35% revenue growth\n• 1,200 new customers\n• 4.8/5 satisfaction rating",
          "x": 0.5,
          "y": 2.0,
          "w": 9.0,
          "h": 3.0,
          "fontSize": 20,
          "bullet": true
        }
      ]
    }
  ]
}
```

## 🛠️ Available Tools

| Tool | Description |
|------|-------------|
| **create_excel** | Create Excel workbooks with sheets, data, formulas, and charts |
| **create_word** | Create Word documents with paragraphs, tables, images, and formatting |
| **create_powerpoint** | Create PowerPoint presentations with slides, text, images, and charts |
| **excel_to_csv** | Convert Excel workbooks to CSV format |

## 📚 Documentation

### Excel Features
- **Data Types**: Numbers, text, dates, formulas, booleans
- **Formulas**: SUM, AVERAGE, COUNT, IF, VLOOKUP, CONCATENATE, and more
- **Formatting**: Fonts, colors, borders, alignment, number formats
- **Charts**: Line, bar, pie, scatter, area with customization
- **Pivot Tables**: Dynamic data analysis and summarization

### Word Features
- **Structure**: Sections, headers, footers, page breaks
- **Text**: Bold, italic, underline, strike, colors, fonts, sizes
- **Paragraphs**: Alignment, spacing, bullets, numbering, headings
- **Tables**: Cells, borders, shading, merging, column/row spanning
- **Special**: Table of contents, images, hyperlinks

### PowerPoint Features
- **Layouts**: Title, content, section, comparison, blank
- **Themes**: Default, light, dark, colorful with custom colors
- **Content**: Text, images, shapes, tables, charts
- **Formatting**: Fonts, colors, alignment, bullets, borders
- **Notes**: Speaker notes for each slide

## 🎯 Use Cases

- **Business Reports** - Automated financial reports with charts and summaries
- **Documentation** - Technical docs, manuals, and guides
- **Presentations** - Sales pitches, project reviews, training materials
- **Data Analysis** - Spreadsheets with complex formulas and pivot tables
- **Templates** - Reusable document templates for teams
- **Batch Processing** - Generate hundreds of personalized documents

## 🧪 Advanced Examples

### Complex Excel with Formulas and Charts

```javascript
{
  "filename": "financial_analysis.xlsx",
  "sheets": [{
    "name": "Revenue",
    "data": [
      ["Product", "Q1", "Q2", "Q3", "Q4", "Total"],
      ["Software", 100000, 120000, 115000, 140000, "=SUM(B2:E2)"],
      ["Services", 50000, 55000, 60000, 65000, "=SUM(B3:E3)"],
      ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)", "=SUM(E2:E3)", "=SUM(F2:F3)"]
    ],
    "charts": [{
      "type": "line",
      "title": "Quarterly Revenue Trend",
      "dataRange": "A1:E3"
    }]
  }]
}
```

### Multi-Section Word Document

```javascript
{
  "filename": "user_manual.docx",
  "sections": [
    {
      "children": [
        { "type": "paragraph", "text": "User Manual", "heading": "Heading1" },
        { "type": "paragraph", "text": "Version 2.0" },
        { "type": "pageBreak" },
        { "type": "paragraph", "text": "Table of Contents", "heading": "Heading1" },
        { "type": "toc" },
        { "type": "pageBreak" },
        { "type": "paragraph", "text": "Introduction", "heading": "Heading1" },
        { "type": "paragraph", "text": "This manual provides comprehensive guidance..." }
      ]
    }
  ]
}
```

## 🤝 Contributing

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🌟 Why Office Whisperer?

- **1.2B Office Users** - Massive market opportunity
- **Natural Language** - No complex APIs or VBA scripting
- **Cross-Platform** - Works on Windows, macOS, and Linux
- **File-Based** - No Office installation required
- **AI-Powered** - Claude understands context and intent
- **Production-Ready** - Battle-tested libraries (ExcelJS, docx, PptxGenJS)

---

**Built with ❤️ using TypeScript and the Model Context Protocol**
