#!/usr/bin/env node
/**
 * Office Whisperer - MCP Server  
 * 38+ Professional Tools for Microsoft Office Suite Automation
 * Control Excel, Word, PowerPoint, and Outlook through natural language with Claude Desktop
 */

import { ExcelGenerator } from './generators/excel-generator.js';
import { WordGenerator } from './generators/word-generator.js';
import { PowerPointGenerator } from './generators/powerpoint-generator.js';
import { OutlookGenerator } from './generators/outlook-generator.js';
import type { MCPRequest, MCPResponse, MCPTool } from './types.js';
import * as fs from 'fs/promises';
import * as path from 'path';

class OfficeWhispererServer {
  private excelGen = new ExcelGenerator();
  private wordGen = new WordGenerator();
  private pptGen = new PowerPointGenerator();
  private outlookGen = new OutlookGenerator();

  private tools: MCPTool[] = [
    // ========== EXCEL TOOLS (15 tools) ==========
    {
      name: 'create_excel',
      description: '📊 Create Excel workbook with sheets, data, formulas, and charts',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string', description: 'Output filename (e.g., "report.xlsx")' },
          sheets: { type: 'array', description: 'Array of sheet configurations' },
          outputPath: { type: 'string', description: 'Optional output directory' },
        },
        required: ['filename', 'sheets'],
      },
    },
    {
      name: 'excel_add_pivot_table',
      description: '📈 Add pivot table with rows, columns, values, and filters',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          pivotTable: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'pivotTable'],
      },
    },
    {
      name: 'excel_add_chart',
      description: '📉 Add chart (line/bar/pie/scatter/area) with customization',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          chart: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'chart'],
      },
    },
    {
      name: 'excel_add_formula',
      description: '🔢 Add formulas (VLOOKUP, SUMIF, INDEX/MATCH, IF, etc)',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          formulas: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'formulas'],
      },
    },
    {
      name: 'excel_conditional_formatting',
      description: '🎨 Apply conditional formatting (color scales, data bars, icon sets)',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          range: { type: 'string' },
          rules: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'range', 'rules'],
      },
    },
    {
      name: 'excel_data_validation',
      description: '✅ Add dropdown lists and validation rules',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          range: { type: 'string' },
          validation: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'range', 'validation'],
      },
    },
    {
      name: 'excel_freeze_panes',
      description: '❄️ Freeze rows/columns for scrolling',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          row: { type: 'number' },
          column: { type: 'number' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName'],
      },
    },
    {
      name: 'excel_filter_sort',
      description: '🔍 Apply AutoFilter and sorting',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          range: { type: 'string' },
          sortBy: { type: 'array' },
          autoFilter: { type: 'boolean' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName'],
      },
    },
    {
      name: 'excel_format_cells',
      description: '✨ Format cells (fonts, colors, borders, alignment, number formats)',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          range: { type: 'string' },
          style: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName', 'range', 'style'],
      },
    },
    {
      name: 'excel_named_range',
      description: '🏷️ Create and manage named ranges',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          name: { type: 'string' },
          range: { type: 'string' },
          sheetName: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'name', 'range'],
      },
    },
    {
      name: 'excel_protect_sheet',
      description: '🔒 Protect worksheets with passwords',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          password: { type: 'string' },
          options: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sheetName'],
      },
    },
    {
      name: 'excel_merge_workbooks',
      description: '🔗 Merge multiple Excel files',
      inputSchema: {
        type: 'object',
        properties: {
          files: { type: 'array' },
          outputFilename: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['files', 'outputFilename'],
      },
    },
    {
      name: 'excel_find_replace',
      description: '🔎 Find and replace values/formulas',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          sheetName: { type: 'string' },
          find: { type: 'string' },
          replace: { type: 'string' },
          matchCase: { type: 'boolean' },
          matchEntireCell: { type: 'boolean' },
          searchFormulas: { type: 'boolean' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'find', 'replace'],
      },
    },
    {
      name: 'excel_to_json',
      description: '📋 Export Excel to JSON format',
      inputSchema: {
        type: 'object',
        properties: {
          excelPath: { type: 'string' },
          sheetName: { type: 'string' },
          outputPath: { type: 'string' },
          header: { type: 'boolean' },
        },
        required: ['excelPath'],
      },
    },
    {
      name: 'excel_to_csv',
      description: '📄 Convert Excel to CSV format',
      inputSchema: {
        type: 'object',
        properties: {
          excelPath: { type: 'string' },
          sheetName: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['excelPath'],
      },
    },

    // ========== WORD TOOLS (10 tools) ==========
    {
      name: 'create_word',
      description: '📝 Create Word document with paragraphs, tables, images, and formatting',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          title: { type: 'string' },
          sections: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'sections'],
      },
    },
    {
      name: 'word_add_toc',
      description: '📑 Add table of contents',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          outputPath: { type: 'string' },
          title: { type: 'string' },
          hyperlinks: { type: 'boolean' },
          levels: { type: 'number' },
        },
        required: ['filename'],
      },
    },
    {
      name: 'word_mail_merge',
      description: '✉️ Mail merge with data source',
      inputSchema: {
        type: 'object',
        properties: {
          templatePath: { type: 'string' },
          dataSource: { type: 'array' },
          outputPath: { type: 'string' },
          outputFilename: { type: 'string' },
        },
        required: ['templatePath', 'dataSource'],
      },
    },
    {
      name: 'word_find_replace',
      description: '🔍 Find and replace text with formatting',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          find: { type: 'string' },
          replace: { type: 'string' },
          matchCase: { type: 'boolean' },
          matchWholeWord: { type: 'boolean' },
          formatting: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'find', 'replace'],
      },
    },
    {
      name: 'word_add_comment',
      description: '💬 Add comments and track changes',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          text: { type: 'string' },
          comment: { type: 'string' },
          author: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'text', 'comment'],
      },
    },
    {
      name: 'word_format_styles',
      description: '🎨 Apply and customize styles/themes',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          styles: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'styles'],
      },
    },
    {
      name: 'word_insert_image',
      description: '🖼️ Insert and position images with wrapping',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          imagePath: { type: 'string' },
          position: { type: 'object' },
          size: { type: 'object' },
          wrapping: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'imagePath'],
      },
    },
    {
      name: 'word_add_header_footer',
      description: '📄 Customize headers/footers per section',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          type: { type: 'string', enum: ['header', 'footer'] },
          content: { type: 'array' },
          sectionType: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'type', 'content'],
      },
    },
    {
      name: 'word_compare_documents',
      description: '🔄 Compare two documents and show differences',
      inputSchema: {
        type: 'object',
        properties: {
          originalPath: { type: 'string' },
          revisedPath: { type: 'string' },
          outputPath: { type: 'string' },
          author: { type: 'string' },
        },
        required: ['originalPath', 'revisedPath'],
      },
    },
    {
      name: 'word_to_pdf',
      description: '📑 Export Word document to PDF',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename'],
      },
    },

    // ========== POWERPOINT TOOLS (8 tools) ==========
    {
      name: 'create_powerpoint',
      description: '🎬 Create PowerPoint presentation with slides, text, images, and charts',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          title: { type: 'string' },
          theme: { type: 'string', enum: ['default', 'light', 'dark', 'colorful'] },
          slides: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slides'],
      },
    },
    {
      name: 'ppt_add_transition',
      description: '✨ Add slide transitions (fade, push, wipe, etc)',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideNumber: { type: 'number' },
          transition: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'transition'],
      },
    },
    {
      name: 'ppt_add_animation',
      description: '🎭 Add animations to objects (entrance, emphasis, exit)',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideNumber: { type: 'number' },
          objectId: { type: 'string' },
          animation: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slideNumber', 'animation'],
      },
    },
    {
      name: 'ppt_add_notes',
      description: '📝 Add/edit speaker notes',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideNumber: { type: 'number' },
          notes: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slideNumber', 'notes'],
      },
    },
    {
      name: 'ppt_duplicate_slide',
      description: '📋 Duplicate slides',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideNumber: { type: 'number' },
          position: { type: 'number' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slideNumber'],
      },
    },
    {
      name: 'ppt_reorder_slides',
      description: '🔀 Reorder slide sequence',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideOrder: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slideOrder'],
      },
    },
    {
      name: 'ppt_export_pdf',
      description: '📑 Export presentation to PDF',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['filename'],
      },
    },
    {
      name: 'ppt_add_media',
      description: '🎥 Embed video/audio',
      inputSchema: {
        type: 'object',
        properties: {
          filename: { type: 'string' },
          slideNumber: { type: 'number' },
          mediaPath: { type: 'string' },
          mediaType: { type: 'string', enum: ['video', 'audio'] },
          position: { type: 'object' },
          size: { type: 'object' },
          outputPath: { type: 'string' },
        },
        required: ['filename', 'slideNumber', 'mediaPath', 'mediaType'],
      },
    },

    // ========== OUTLOOK TOOLS (5 tools) ==========
    {
      name: 'outlook_send_email',
      description: '📧 Send emails with attachments',
      inputSchema: {
        type: 'object',
        properties: {
          to: { type: 'string' },
          subject: { type: 'string' },
          body: { type: 'string' },
          cc: { type: 'string' },
          bcc: { type: 'string' },
          attachments: { type: 'array' },
          html: { type: 'boolean' },
          priority: { type: 'string', enum: ['high', 'normal', 'low'] },
          smtpConfig: { type: 'object' },
        },
        required: ['to', 'subject', 'body'],
      },
    },
    {
      name: 'outlook_create_meeting',
      description: '📅 Create calendar events with attendees',
      inputSchema: {
        type: 'object',
        properties: {
          subject: { type: 'string' },
          startTime: { type: 'string' },
          endTime: { type: 'string' },
          location: { type: 'string' },
          attendees: { type: 'array' },
          description: { type: 'string' },
          reminder: { type: 'number' },
          outputPath: { type: 'string' },
        },
        required: ['subject', 'startTime', 'endTime'],
      },
    },
    {
      name: 'outlook_add_contact',
      description: '👤 Add contact to address book',
      inputSchema: {
        type: 'object',
        properties: {
          firstName: { type: 'string' },
          lastName: { type: 'string' },
          email: { type: 'string' },
          phone: { type: 'string' },
          company: { type: 'string' },
          jobTitle: { type: 'string' },
          address: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['firstName', 'lastName'],
      },
    },
    {
      name: 'outlook_create_task',
      description: '✅ Create Outlook task',
      inputSchema: {
        type: 'object',
        properties: {
          subject: { type: 'string' },
          dueDate: { type: 'string' },
          priority: { type: 'string', enum: ['high', 'normal', 'low'] },
          status: { type: 'string' },
          category: { type: 'string' },
          reminder: { type: 'string' },
          notes: { type: 'string' },
          outputPath: { type: 'string' },
        },
        required: ['subject'],
      },
    },
    {
      name: 'outlook_set_rule',
      description: '⚙️ Create inbox rule',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string' },
          conditions: { type: 'array' },
          actions: { type: 'array' },
          outputPath: { type: 'string' },
        },
        required: ['name', 'conditions', 'actions'],
      },
    },
  ];

  async start(): Promise<void> {
    process.stdin.setEncoding('utf-8');
    let buffer = '';

    process.stdin.on('data', async (chunk) => {
      buffer += chunk;
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        if (line.trim()) {
          try {
            const request: MCPRequest = JSON.parse(line);
            const response = await this.handleRequest(request);
            if (response) {
              console.log(JSON.stringify(response));
            }
          } catch (error) {
            console.error('Error processing request:', error);
          }
        }
      }
    });

    process.stdin.on('end', () => {
      process.exit(0);
    });
  }

  private async handleRequest(request: MCPRequest): Promise<MCPResponse | null> {
    const { id, method, params } = request;

    if (id === undefined) {
      if (method === 'notifications/initialized') {
        return null;
      }
      return null;
    }

    try {
      switch (method) {
        case 'initialize':
          return this.initialize(id);
        case 'tools/list':
          return {
            jsonrpc: '2.0',
            id,
            result: { tools: this.tools },
          };
        case 'tools/call':
          return await this.callTool(id, params);
        default:
          return {
            jsonrpc: '2.0',
            id,
            error: {
              code: -32601,
              message: `Method not found: ${method}`,
            },
          };
      }
    } catch (error) {
      return {
        jsonrpc: '2.0',
        id,
        error: {
          code: -32603,
          message: error instanceof Error ? error.message : 'Internal error',
        },
      };
    }
  }

  private initialize(id: string | number): MCPResponse {
    return {
      jsonrpc: '2.0',
      id,
      result: {
        protocolVersion: '2025-06-18',
        capabilities: {
          tools: {},
        },
        serverInfo: {
          name: 'office-whisperer',
          version: '2.0.0',
        },
      },
    };
  }

  private async callTool(id: string | number, params: any): Promise<MCPResponse> {
    const { name, arguments: args } = params;
    let result: any;

    try {
      // Excel Tools
      if (name === 'create_excel') {
        result = await this.handleCreateExcel(args);
      } else if (name === 'excel_add_pivot_table') {
        result = await this.handleExcelAddPivotTable(args);
      } else if (name === 'excel_add_chart') {
        result = await this.handleExcelAddChart(args);
      } else if (name === 'excel_add_formula') {
        result = await this.handleExcelAddFormula(args);
      } else if (name === 'excel_conditional_formatting') {
        result = await this.handleExcelConditionalFormatting(args);
      } else if (name === 'excel_data_validation') {
        result = await this.handleExcelDataValidation(args);
      } else if (name === 'excel_freeze_panes') {
        result = await this.handleExcelFreezePanes(args);
      } else if (name === 'excel_filter_sort') {
        result = await this.handleExcelFilterSort(args);
      } else if (name === 'excel_format_cells') {
        result = await this.handleExcelFormatCells(args);
      } else if (name === 'excel_named_range') {
        result = await this.handleExcelNamedRange(args);
      } else if (name === 'excel_protect_sheet') {
        result = await this.handleExcelProtectSheet(args);
      } else if (name === 'excel_merge_workbooks') {
        result = await this.handleExcelMergeWorkbooks(args);
      } else if (name === 'excel_find_replace') {
        result = await this.handleExcelFindReplace(args);
      } else if (name === 'excel_to_json') {
        result = await this.handleExcelToJSON(args);
      } else if (name === 'excel_to_csv') {
        result = await this.handleExcelToCSV(args);
      }
      
      // Word Tools
      else if (name === 'create_word') {
        result = await this.handleCreateWord(args);
      } else if (name === 'word_add_toc') {
        result = await this.handleWordAddTOC(args);
      } else if (name === 'word_mail_merge') {
        result = await this.handleWordMailMerge(args);
      } else if (name === 'word_find_replace') {
        result = await this.handleWordFindReplace(args);
      } else if (name === 'word_add_comment') {
        result = await this.handleWordAddComment(args);
      } else if (name === 'word_format_styles') {
        result = await this.handleWordFormatStyles(args);
      } else if (name === 'word_insert_image') {
        result = await this.handleWordInsertImage(args);
      } else if (name === 'word_add_header_footer') {
        result = await this.handleWordAddHeaderFooter(args);
      } else if (name === 'word_compare_documents') {
        result = await this.handleWordCompareDocuments(args);
      } else if (name === 'word_to_pdf') {
        result = await this.handleWordToPDF(args);
      }
      
      // PowerPoint Tools
      else if (name === 'create_powerpoint') {
        result = await this.handleCreatePowerPoint(args);
      } else if (name === 'ppt_add_transition') {
        result = await this.handlePPTAddTransition(args);
      } else if (name === 'ppt_add_animation') {
        result = await this.handlePPTAddAnimation(args);
      } else if (name === 'ppt_add_notes') {
        result = await this.handlePPTAddNotes(args);
      } else if (name === 'ppt_duplicate_slide') {
        result = await this.handlePPTDuplicateSlide(args);
      } else if (name === 'ppt_reorder_slides') {
        result = await this.handlePPTReorderSlides(args);
      } else if (name === 'ppt_export_pdf') {
        result = await this.handlePPTExportPDF(args);
      } else if (name === 'ppt_add_media') {
        result = await this.handlePPTAddMedia(args);
      }
      
      // Outlook Tools
      else if (name === 'outlook_send_email') {
        result = await this.handleOutlookSendEmail(args);
      } else if (name === 'outlook_create_meeting') {
        result = await this.handleOutlookCreateMeeting(args);
      } else if (name === 'outlook_add_contact') {
        result = await this.handleOutlookAddContact(args);
      } else if (name === 'outlook_create_task') {
        result = await this.handleOutlookCreateTask(args);
      } else if (name === 'outlook_set_rule') {
        result = await this.handleOutlookSetRule(args);
      }
      
      else {
        throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      throw error;
    }

    return {
      jsonrpc: '2.0',
      id,
      result: {
        content: [
          {
            type: 'text',
            text: result,
          },
        ],
      },
    };
  }

  // ========== EXCEL HANDLERS ==========
  
  private async handleCreateExcel(args: any): Promise<string> {
    const buffer = await this.excelGen.createWorkbook({
      filename: args.filename,
      sheets: args.sheets,
    });
    const outputPath = args.outputPath || process.cwd();
    const fullPath = path.join(outputPath, args.filename);
    await fs.writeFile(fullPath, buffer);
    return `✅ **Excel workbook created!**\n\n📊 **File:** ${fullPath}\n📝 **Sheets:** ${args.sheets.length}\n💾 **Size:** ${(buffer.length / 1024).toFixed(2)} KB`;
  }

  private async handleExcelAddPivotTable(args: any): Promise<string> {
    const buffer = await this.excelGen.addPivotTable(args.filename, args.sheetName, args.pivotTable);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Pivot table added!**\n\n📈 **Name:** ${args.pivotTable.name}\n📊 **File:** ${fullPath}`;
  }

  private async handleExcelAddChart(args: any): Promise<string> {
    const buffer = await this.excelGen.addChart(args.filename, args.sheetName, args.chart);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Chart added!**\n\n📉 **Type:** ${args.chart.type}\n📊 **Title:** ${args.chart.title}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelAddFormula(args: any): Promise<string> {
    const buffer = await this.excelGen.addFormulas(args.filename, args.sheetName, args.formulas);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Formulas added!**\n\n🔢 **Count:** ${args.formulas.length}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelConditionalFormatting(args: any): Promise<string> {
    const buffer = await this.excelGen.addConditionalFormatting(args.filename, args.sheetName, args.range, args.rules);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Conditional formatting applied!**\n\n🎨 **Range:** ${args.range}\n📊 **Rules:** ${args.rules.length}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelDataValidation(args: any): Promise<string> {
    const buffer = await this.excelGen.addDataValidation(args.filename, args.sheetName, args.range, args.validation);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Data validation added!**\n\n✅ **Type:** ${args.validation.type}\n📍 **Range:** ${args.range}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelFreezePanes(args: any): Promise<string> {
    const buffer = await this.excelGen.freezePanes(args.filename, args.sheetName, args.row, args.column);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Panes frozen!**\n\n❄️ **Row:** ${args.row || 0}\n❄️ **Column:** ${args.column || 0}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelFilterSort(args: any): Promise<string> {
    const buffer = await this.excelGen.filterSort(args.filename, args.sheetName, args.range, args.sortBy, args.autoFilter);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Filter/Sort applied!**\n\n🔍 **AutoFilter:** ${args.autoFilter ? 'Yes' : 'No'}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelFormatCells(args: any): Promise<string> {
    const buffer = await this.excelGen.formatCells(args.filename, args.sheetName, args.range, args.style);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Cells formatted!**\n\n✨ **Range:** ${args.range}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelNamedRange(args: any): Promise<string> {
    const buffer = await this.excelGen.addNamedRange(args.filename, args.name, args.range, args.sheetName);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Named range created!**\n\n🏷️ **Name:** ${args.name}\n📍 **Range:** ${args.range}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelProtectSheet(args: any): Promise<string> {
    const buffer = await this.excelGen.protectSheet(args.filename, args.sheetName, args.password, args.options);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Sheet protected!**\n\n🔒 **Sheet:** ${args.sheetName}\n${args.password ? '🔑 **Password:** Set' : '🔓 **Password:** None'}\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelMergeWorkbooks(args: any): Promise<string> {
    const buffer = await this.excelGen.mergeWorkbooks(args.files, args.outputFilename);
    const outputPath = args.outputPath || process.cwd();
    const fullPath = path.join(outputPath, args.outputFilename);
    await fs.writeFile(fullPath, buffer);
    return `✅ **Workbooks merged!**\n\n🔗 **Source files:** ${args.files.length}\n📁 **Output:** ${fullPath}`;
  }

  private async handleExcelFindReplace(args: any): Promise<string> {
    const buffer = await this.excelGen.findReplace(
      args.filename,
      args.find,
      args.replace,
      args.sheetName,
      args.matchCase,
      args.matchEntireCell,
      args.searchFormulas
    );
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Find & Replace complete!**\n\n🔎 **Find:** "${args.find}"\n🔄 **Replace:** "${args.replace}"\n📁 **File:** ${fullPath}`;
  }

  private async handleExcelToJSON(args: any): Promise<string> {
    const json = await this.excelGen.convertToJSON(args.excelPath, args.sheetName, args.header);
    const outputPath = args.outputPath || args.excelPath.replace(/\.xlsx?$/i, '.json');
    await fs.writeFile(outputPath, json, 'utf-8');
    return `✅ **Converted to JSON!**\n\n📊 **Source:** ${args.excelPath}\n📋 **Output:** ${outputPath}\n💾 **Size:** ${(json.length / 1024).toFixed(2)} KB`;
  }

  private async handleExcelToCSV(args: any): Promise<string> {
    const csv = await this.excelGen.convertToCSV(args.excelPath, args.sheetName);
    const outputPath = args.outputPath || args.excelPath.replace(/\.xlsx?$/i, '.csv');
    await fs.writeFile(outputPath, csv, 'utf-8');
    return `✅ **Converted to CSV!**\n\n📊 **Source:** ${args.excelPath}\n📝 **Output:** ${outputPath}\n💾 **Size:** ${(csv.length / 1024).toFixed(2)} KB`;
  }

  // ========== WORD HANDLERS ==========
  
  private async handleCreateWord(args: any): Promise<string> {
    const buffer = await this.wordGen.createDocument({
      filename: args.filename,
      sections: args.sections,
    });
    const outputPath = args.outputPath || process.cwd();
    const fullPath = path.join(outputPath, args.filename);
    await fs.writeFile(fullPath, buffer);
    return `✅ **Word document created!**\n\n📄 **File:** ${fullPath}\n📑 **Sections:** ${args.sections.length}\n💾 **Size:** ${(buffer.length / 1024).toFixed(2)} KB`;
  }

  private async handleWordAddTOC(args: any): Promise<string> {
    const buffer = await this.wordGen.addTableOfContents(args.filename, args.title, args.hyperlinks, args.levels);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Table of Contents added!**\n\n📑 **Title:** ${args.title || 'Table of Contents'}\n📁 **File:** ${fullPath}`;
  }

  private async handleWordMailMerge(args: any): Promise<string> {
    const buffers = await this.wordGen.mailMerge(args.templatePath, args.dataSource, args.outputFilename);
    const outputPath = args.outputPath || process.cwd();
    let savedFiles = 0;
    for (let i = 0; i < buffers.length; i++) {
      const filename = args.outputFilename ? `${args.outputFilename}_${i + 1}.docx` : `merged_${i + 1}.docx`;
      const fullPath = path.join(outputPath, filename);
      await fs.writeFile(fullPath, buffers[i]);
      savedFiles++;
    }
    return `✅ **Mail merge complete!**\n\n✉️ **Documents created:** ${savedFiles}\n📁 **Output:** ${outputPath}`;
  }

  private async handleWordFindReplace(args: any): Promise<string> {
    const buffer = await this.wordGen.findReplace(
      args.filename,
      args.find,
      args.replace,
      args.matchCase,
      args.matchWholeWord,
      args.formatting
    );
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Find & Replace complete!**\n\n🔍 **Find:** "${args.find}"\n🔄 **Replace:** "${args.replace}"\n📁 **File:** ${fullPath}`;
  }

  private async handleWordAddComment(args: any): Promise<string> {
    const buffer = await this.wordGen.addComment(args.filename, args.text, args.comment, args.author);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Comment added!**\n\n💬 **Author:** ${args.author || 'Office Whisperer'}\n📁 **File:** ${fullPath}`;
  }

  private async handleWordFormatStyles(args: any): Promise<string> {
    const buffer = await this.wordGen.formatStyles(args.filename, args.styles);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Styles applied!**\n\n🎨 **Custom styles added**\n📁 **File:** ${fullPath}`;
  }

  private async handleWordInsertImage(args: any): Promise<string> {
    const buffer = await this.wordGen.insertImage(
      args.filename,
      args.imagePath,
      args.position,
      args.size,
      args.wrapping
    );
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Image inserted!**\n\n🖼️ **Source:** ${args.imagePath}\n📁 **File:** ${fullPath}`;
  }

  private async handleWordAddHeaderFooter(args: any): Promise<string> {
    const buffer = await this.wordGen.addHeaderFooter(
      args.filename,
      args.type,
      args.content,
      args.sectionType
    );
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **${args.type.charAt(0).toUpperCase() + args.type.slice(1)} added!**\n\n📄 **Type:** ${args.type}\n📁 **File:** ${fullPath}`;
  }

  private async handleWordCompareDocuments(args: any): Promise<string> {
    const buffer = await this.wordGen.compareDocuments(args.originalPath, args.revisedPath, args.author);
    const outputPath = args.outputPath || path.dirname(args.originalPath);
    const fullPath = path.join(outputPath, 'comparison.docx');
    await fs.writeFile(fullPath, buffer);
    return `✅ **Documents compared!**\n\n🔄 **Original:** ${args.originalPath}\n🔄 **Revised:** ${args.revisedPath}\n📁 **Report:** ${fullPath}`;
  }

  private async handleWordToPDF(args: any): Promise<string> {
    const buffer = await this.wordGen.convertToPDF(args.filename);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const pdfPath = path.join(outputPath, path.basename(args.filename).replace(/\.docx?$/i, '.pdf'));
    await fs.writeFile(pdfPath, buffer);
    return `✅ **PDF conversion info!**\n\n📑 **Source:** ${args.filename}\n📄 **Output:** ${pdfPath}\n\nNote: Actual PDF conversion requires LibreOffice or similar tools.`;
  }

  // ========== POWERPOINT HANDLERS ==========
  
  private async handleCreatePowerPoint(args: any): Promise<string> {
    const buffer = await this.pptGen.createPresentation({
      filename: args.filename,
      title: args.title,
      theme: args.theme,
      slides: args.slides,
    });
    const outputPath = args.outputPath || process.cwd();
    const fullPath = path.join(outputPath, args.filename);
    await fs.writeFile(fullPath, buffer);
    return `✅ **PowerPoint created!**\n\n🎬 **File:** ${fullPath}\n🎨 **Theme:** ${args.theme || 'default'}\n📊 **Slides:** ${args.slides.length}\n💾 **Size:** ${(buffer.length / 1024).toFixed(2)} KB`;
  }

  private async handlePPTAddTransition(args: any): Promise<string> {
    const buffer = await this.pptGen.addTransition(args.filename, args.transition, args.slideNumber);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Transition added!**\n\n✨ **Type:** ${args.transition.type}\n📁 **File:** ${fullPath}`;
  }

  private async handlePPTAddAnimation(args: any): Promise<string> {
    const buffer = await this.pptGen.addAnimation(args.filename, args.slideNumber, args.animation, args.objectId);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Animation added!**\n\n🎭 **Effect:** ${args.animation.effect}\n📊 **Slide:** ${args.slideNumber}\n📁 **File:** ${fullPath}`;
  }

  private async handlePPTAddNotes(args: any): Promise<string> {
    const buffer = await this.pptGen.addNotes(args.filename, args.slideNumber, args.notes);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Speaker notes added!**\n\n📝 **Slide:** ${args.slideNumber}\n📁 **File:** ${fullPath}`;
  }

  private async handlePPTDuplicateSlide(args: any): Promise<string> {
    const buffer = await this.pptGen.duplicateSlide(args.filename, args.slideNumber, args.position);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Slide duplicated!**\n\n📋 **Source:** Slide ${args.slideNumber}\n📍 **Position:** ${args.position || 'end'}\n📁 **File:** ${fullPath}`;
  }

  private async handlePPTReorderSlides(args: any): Promise<string> {
    const buffer = await this.pptGen.reorderSlides(args.filename, args.slideOrder);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Slides reordered!**\n\n🔀 **New order:** ${args.slideOrder.join(', ')}\n📁 **File:** ${fullPath}`;
  }

  private async handlePPTExportPDF(args: any): Promise<string> {
    const buffer = await this.pptGen.exportPDF(args.filename);
    const outputPath = args.outputPath || path.dirname(args.filename);
    const pdfPath = path.join(outputPath, path.basename(args.filename).replace(/\.pptx?$/i, '.pdf'));
    await fs.writeFile(pdfPath, buffer);
    return `✅ **PDF export info!**\n\n📑 **Source:** ${args.filename}\n📄 **Output:** ${pdfPath}\n\nNote: Actual PDF conversion requires LibreOffice or PowerPoint.`;
  }

  private async handlePPTAddMedia(args: any): Promise<string> {
    const buffer = await this.pptGen.addMedia(
      args.filename,
      args.slideNumber,
      args.mediaPath,
      args.mediaType,
      args.position,
      args.size
    );
    const outputPath = args.outputPath || path.dirname(args.filename);
    const fullPath = path.join(outputPath, path.basename(args.filename));
    await fs.writeFile(fullPath, buffer);
    return `✅ **Media embedded!**\n\n🎥 **Type:** ${args.mediaType}\n📊 **Slide:** ${args.slideNumber}\n📁 **File:** ${fullPath}`;
  }

  // ========== OUTLOOK HANDLERS ==========
  
  private async handleOutlookSendEmail(args: any): Promise<string> {
    const result = await this.outlookGen.sendEmail(args);
    return `✅ **Email processed!**\n\n${result}`;
  }

  private async handleOutlookCreateMeeting(args: any): Promise<string> {
    const ics = await this.outlookGen.createMeeting(args);
    const outputPath = args.outputPath || process.cwd();
    const filename = `meeting_${Date.now()}.ics`;
    const fullPath = path.join(outputPath, filename);
    await fs.writeFile(fullPath, ics, 'utf-8');
    return `✅ **Meeting created!**\n\n📅 **Subject:** ${args.subject}\n⏰ **Start:** ${args.startTime}\n📁 **ICS file:** ${fullPath}\n\nImport this file into Outlook/Google Calendar.`;
  }

  private async handleOutlookAddContact(args: any): Promise<string> {
    const vcf = await this.outlookGen.addContact(args);
    const outputPath = args.outputPath || process.cwd();
    const filename = `contact_${args.lastName}_${args.firstName}.vcf`;
    const fullPath = path.join(outputPath, filename);
    await fs.writeFile(fullPath, vcf, 'utf-8');
    return `✅ **Contact created!**\n\n👤 **Name:** ${args.firstName} ${args.lastName}\n${args.email ? `📧 **Email:** ${args.email}\n` : ''}📁 **VCF file:** ${fullPath}\n\nImport this file into Outlook/Contacts app.`;
  }

  private async handleOutlookCreateTask(args: any): Promise<string> {
    const task = await this.outlookGen.createTask(args);
    const outputPath = args.outputPath || process.cwd();
    const filename = `task_${Date.now()}.json`;
    const fullPath = path.join(outputPath, filename);
    await fs.writeFile(fullPath, task, 'utf-8');
    return `✅ **Task created!**\n\n✅ **Subject:** ${args.subject}\n${args.dueDate ? `📅 **Due:** ${args.dueDate}\n` : ''}🔢 **Priority:** ${args.priority || 'normal'}\n📁 **JSON file:** ${fullPath}`;
  }

  private async handleOutlookSetRule(args: any): Promise<string> {
    const rule = await this.outlookGen.setRule(args);
    const outputPath = args.outputPath || process.cwd();
    const filename = `rule_${Date.now()}.json`;
    const fullPath = path.join(outputPath, filename);
    await fs.writeFile(fullPath, rule, 'utf-8');
    return `✅ **Inbox rule created!**\n\n⚙️ **Name:** ${args.name}\n🔧 **Conditions:** ${args.conditions.length}\n🎯 **Actions:** ${args.actions.length}\n📁 **JSON file:** ${fullPath}`;
  }
}

// Start the server
const server = new OfficeWhispererServer();
server.start().catch(console.error);
