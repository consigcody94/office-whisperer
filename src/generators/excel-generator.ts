/**
 * Excel Generator - Create and manipulate Excel workbooks using ExcelJS
 */

import ExcelJS from 'exceljs';
import * as fs from 'fs/promises';
import type {
  ExcelWorkbookOptions,
  ExcelSheet,
  ExcelFormula,
  ExcelChart,
  ExcelPivotTable,
  ExcelConditionalFormattingRule,
  ExcelDataValidation,
  ExcelCellStyle,
} from '../types.js';

export class ExcelGenerator {
  async createWorkbook(options: ExcelWorkbookOptions): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();

    // Set workbook properties
    workbook.creator = 'Office Whisperer';
    workbook.created = new Date();
    workbook.modified = new Date();

    // Create sheets
    for (const sheetConfig of options.sheets) {
      const worksheet = workbook.addWorksheet(sheetConfig.name);

      // Add columns if specified
      if (sheetConfig.columns) {
        worksheet.columns = sheetConfig.columns.map(col => ({
          header: col.header,
          key: col.key,
          width: col.width || 15,
        }));
      }

      // Add data rows
      if (sheetConfig.data) {
        sheetConfig.data.forEach((row, index) => {
          worksheet.addRow(row);
        });
      }

      // Apply row styles
      if (sheetConfig.rows) {
        sheetConfig.rows.forEach((rowConfig, index) => {
          const row = worksheet.getRow(index + 1);
          row.values = rowConfig.values;
          if (rowConfig.style) {
            this.applyRowStyle(row, rowConfig.style);
          }
        });
      }

      // Add header styling
      if (worksheet.columns.length > 0) {
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true, size: 12 };
        headerRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF4472C4' },
        };
        headerRow.font = { ...headerRow.font, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;
      }

      // Auto-filter if data exists
      if (sheetConfig.data && sheetConfig.data.length > 0) {
        worksheet.autoFilter = {
          from: { row: 1, column: 1 },
          to: { row: 1, column: sheetConfig.data[0].length },
        };
      }
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addPivotTable(
    filename: string,
    sheetName: string,
    pivotTable: ExcelPivotTable
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Note: ExcelJS has limited pivot table support
    // This creates a placeholder comment indicating where the pivot table would be
    const pivotSheet = workbook.addWorksheet(pivotTable.name);
    pivotSheet.getCell('A1').value = `Pivot Table: ${pivotTable.name}`;
    pivotSheet.getCell('A2').value = `Data Range: ${pivotTable.dataRange}`;
    pivotSheet.getCell('A3').value = `Rows: ${pivotTable.rows.join(', ')}`;
    pivotSheet.getCell('A4').value = `Columns: ${pivotTable.columns.join(', ')}`;
    pivotSheet.getCell('A5').value = `Values: ${pivotTable.values.join(', ')}`;

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addChart(
    filename: string,
    sheetName: string,
    chart: ExcelChart
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Note: ExcelJS has limited chart support
    // This creates a placeholder indicating chart metadata
    const chartCell = worksheet.getCell(chart.position?.row || 10, chart.position?.col || 1);
    chartCell.value = `[Chart: ${chart.title}]`;
    chartCell.note = `Type: ${chart.type}, Data: ${chart.dataRange}`;
    chartCell.font = { bold: true, color: { argb: 'FF0000FF' } };

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addFormulas(
    filename: string,
    sheetName: string,
    formulas: ExcelFormula[]
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    formulas.forEach(({ cell, formula }) => {
      const cellObj = worksheet.getCell(cell);
      cellObj.value = { formula };
    });

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addConditionalFormatting(
    filename: string,
    sheetName: string,
    range: string,
    rules: ExcelConditionalFormattingRule[]
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    rules.forEach((rule) => {
      const cfRule: any = {
        type: rule.type,
        priority: rule.priority || 1,
      };

      switch (rule.type) {
        case 'colorScale':
          if (rule.gradient) {
            cfRule.cfvo = [
              { type: 'min', value: undefined },
              { type: 'max', value: undefined },
            ];
            cfRule.color = [
              { argb: this.normalizeColor(rule.gradient.start) },
              { argb: this.normalizeColor(rule.gradient.end) },
            ];
          }
          break;
        case 'dataBar':
          cfRule.cfvo = [
            { type: 'min', value: undefined },
            { type: 'max', value: undefined },
          ];
          cfRule.color = rule.color ? { argb: this.normalizeColor(rule.color) } : { argb: 'FF638EC6' };
          break;
        case 'iconSet':
          cfRule.iconSet = rule.iconSet || 'ThreeArrows';
          break;
        case 'formulaBased':
          cfRule.formulae = [rule.formula];
          cfRule.style = {
            fill: {
              type: 'pattern',
              pattern: 'solid',
              bgColor: { argb: this.normalizeColor(rule.color || 'FFFF0000') },
            },
          };
          break;
        case 'cellValue':
          cfRule.operator = rule.operator || 'greaterThan';
          cfRule.formulae = rule.values || [];
          cfRule.style = {
            fill: {
              type: 'pattern',
              pattern: 'solid',
              bgColor: { argb: this.normalizeColor(rule.color || 'FFFF0000') },
            },
          };
          break;
      }

      worksheet.addConditionalFormatting({
        ref: range,
        rules: [cfRule],
      });
    });

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addDataValidation(
    filename: string,
    sheetName: string,
    range: string,
    validation: ExcelDataValidation
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    const validationRule: any = {
      type: validation.type,
      allowBlank: validation.allowBlank !== false,
      showErrorMessage: validation.showErrorMessage !== false,
      showInputMessage: validation.showInputMessage !== false,
    };

    if (validation.type === 'list' && validation.values) {
      validationRule.formulae = [`"${validation.values.join(',')}"`];
    } else if (validation.formula) {
      validationRule.formulae = [validation.formula];
    } else if (validation.operator) {
      validationRule.operator = validation.operator;
      validationRule.formulae = [validation.min, validation.max].filter(v => v !== undefined);
    }

    if (validation.errorTitle) {
      validationRule.errorTitle = validation.errorTitle;
      validationRule.error = validation.error || 'Invalid value';
    }

    if (validation.promptTitle) {
      validationRule.promptTitle = validation.promptTitle;
      validationRule.prompt = validation.prompt || '';
    }

    // Apply to range
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (rangeMatch) {
      const [, startCol, startRow, endCol, endRow] = rangeMatch;
      const startColNum = this.columnToNumber(startCol);
      const endColNum = this.columnToNumber(endCol);

      for (let row = parseInt(startRow); row <= parseInt(endRow); row++) {
        for (let col = startColNum; col <= endColNum; col++) {
          worksheet.getCell(row, col).dataValidation = validationRule;
        }
      }
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async freezePanes(
    filename: string,
    sheetName: string,
    row?: number,
    column?: number
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    if (row && column) {
      worksheet.views = [
        { state: 'frozen', xSplit: column, ySplit: row },
      ];
    } else if (row) {
      worksheet.views = [
        { state: 'frozen', ySplit: row },
      ];
    } else if (column) {
      worksheet.views = [
        { state: 'frozen', xSplit: column },
      ];
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async filterSort(
    filename: string,
    sheetName: string,
    range?: string,
    sortBy?: { column: string | number; descending?: boolean }[],
    autoFilter: boolean = true
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    if (autoFilter) {
      if (range) {
        const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (rangeMatch) {
          worksheet.autoFilter = range;
        }
      } else {
        // Auto-detect range
        const dimensions = worksheet.dimensions;
        if (dimensions) {
          worksheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1, column: dimensions.right },
          };
        }
      }
    }

    // Note: ExcelJS doesn't support programmatic sorting, but we can mark it
    if (sortBy) {
      const noteCell = worksheet.getCell('A1');
      noteCell.note = `Sort by: ${sortBy.map(s => `${s.column} ${s.descending ? 'DESC' : 'ASC'}`).join(', ')}`;
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async formatCells(
    filename: string,
    sheetName: string,
    range: string,
    style: ExcelCellStyle
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (rangeMatch) {
      const [, startCol, startRow, endCol, endRow] = rangeMatch;
      const startColNum = this.columnToNumber(startCol);
      const endColNum = this.columnToNumber(endCol);

      for (let row = parseInt(startRow); row <= parseInt(endRow); row++) {
        for (let col = startColNum; col <= endColNum; col++) {
          const cell = worksheet.getCell(row, col);
          if (style.font) cell.font = style.font as any;
          if (style.fill) cell.fill = style.fill as any;
          if (style.alignment) cell.alignment = style.alignment;
          if (style.border) cell.border = style.border as any;
          if (style.numFmt) cell.numFmt = style.numFmt;
        }
      }
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async addNamedRange(
    filename: string,
    name: string,
    range: string,
    sheetName?: string
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);

    workbook.definedNames.add(
      sheetName ? `${sheetName}!${range}` : range,
      name
    );

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async protectSheet(
    filename: string,
    sheetName: string,
    password?: string,
    options?: any
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    worksheet.protect(password || '', options || {});

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async mergeWorkbooks(files: string[], outputFilename: string): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();

    for (const file of files) {
      try {
        const sourceWorkbook = new ExcelJS.Workbook();
        await sourceWorkbook.xlsx.readFile(file);

        sourceWorkbook.eachSheet((worksheet, sheetId) => {
          const newSheet = workbook.addWorksheet(worksheet.name);

          // Copy data
          worksheet.eachRow((row, rowNumber) => {
            const newRow = newSheet.getRow(rowNumber);
            newRow.values = row.values;
            newRow.height = row.height;

            // Copy cell styles
            row.eachCell((cell, colNumber) => {
              const newCell = newRow.getCell(colNumber);
              newCell.style = cell.style;
            });
          });

          // Copy column widths
          worksheet.columns.forEach((column, idx) => {
            if (newSheet.columns[idx]) {
              newSheet.columns[idx].width = column.width;
            }
          });
        });
      } catch (error) {
        console.error(`Error merging file ${file}:`, error);
      }
    }

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async findReplace(
    filename: string,
    find: string,
    replace: string,
    sheetName?: string,
    matchCase: boolean = false,
    matchEntireCell: boolean = false,
    searchFormulas: boolean = false
  ): Promise<Buffer> {
    const workbook = await this.loadWorkbook(filename);

    const sheets = sheetName
      ? [workbook.getWorksheet(sheetName)]
      : workbook.worksheets;

    sheets.forEach((worksheet) => {
      if (!worksheet) return;

      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          let cellValue = searchFormulas && cell.formula ? cell.formula : cell.value;

          if (typeof cellValue === 'string') {
            const searchValue = matchCase ? find : find.toLowerCase();
            const compareValue = matchCase ? cellValue : cellValue.toLowerCase();

            if (matchEntireCell) {
              if (compareValue === searchValue) {
                cell.value = replace;
              }
            } else {
              const regex = new RegExp(
                find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                matchCase ? 'g' : 'gi'
              );
              cell.value = cellValue.replace(regex, replace);
            }
          }
        });
      });
    });

    return await workbook.xlsx.writeBuffer() as unknown as Buffer;
  }

  async convertToJSON(
    excelPath: string,
    sheetName?: string,
    header: boolean = true
  ): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const worksheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[0];

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName || 'default'}" not found`);
    }

    const rows: any[] = [];
    let headers: string[] = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1 && header) {
        headers = row.values as string[];
        headers.shift(); // Remove index 0
      } else {
        const values = row.values as any[];
        values.shift(); // Remove index 0

        if (header && headers.length > 0) {
          const obj: any = {};
          headers.forEach((header, idx) => {
            obj[header] = values[idx];
          });
          rows.push(obj);
        } else {
          rows.push(values);
        }
      }
    });

    return JSON.stringify(rows, null, 2);
  }

  async convertToCSV(excelPath: string, sheetName?: string): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const worksheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[0];

    if (!worksheet) {
      throw new Error(`Sheet "${sheetName || 'default'}" not found`);
    }

    let csv = '';
    worksheet.eachRow((row) => {
      const values = row.values as any[];
      csv += values.slice(1).map(v => {
        // Escape CSV values that contain commas or quotes
        const str = String(v === null || v === undefined ? '' : v);
        return str.includes(',') || str.includes('"') || str.includes('\n')
          ? `"${str.replace(/"/g, '""')}"`
          : str;
      }).join(',') + '\n';
    });

    return csv;
  }

  // Helper methods
  private async loadWorkbook(filename: string): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(filename);
    } catch (error) {
      // If file doesn't exist, return empty workbook
      console.warn(`File ${filename} not found, creating new workbook`);
    }
    return workbook;
  }

  private applyRowStyle(row: ExcelJS.Row, style: any): void {
    if (style.font) {
      row.font = style.font;
    }
    if (style.fill) {
      row.fill = style.fill;
    }
    if (style.alignment) {
      row.alignment = style.alignment;
    }
    if (style.border) {
      row.border = style.border;
    }
  }

  private normalizeColor(color: string): string {
    // Ensure color is in ARGB format
    if (color.startsWith('#')) {
      return 'FF' + color.slice(1).toUpperCase();
    }
    if (color.length === 6) {
      return 'FF' + color.toUpperCase();
    }
    return color.toUpperCase();
  }

  private columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return result;
  }
}
