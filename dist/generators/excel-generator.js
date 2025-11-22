/**
 * Excel Generator - Create and manipulate Excel workbooks using ExcelJS
 */
import ExcelJS from 'exceljs';
export class ExcelGenerator {
    async createWorkbook(options) {
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
        return await workbook.xlsx.writeBuffer();
    }
    async addPivotTable(filename, sheetName, pivotTable) {
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
        return await workbook.xlsx.writeBuffer();
    }
    async addChart(filename, sheetName, chart) {
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
        return await workbook.xlsx.writeBuffer();
    }
    async addFormulas(filename, sheetName, formulas) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        formulas.forEach(({ cell, formula }) => {
            const cellObj = worksheet.getCell(cell);
            cellObj.value = { formula };
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addConditionalFormatting(filename, sheetName, range, rules) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        rules.forEach((rule) => {
            const cfRule = {
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
        return await workbook.xlsx.writeBuffer();
    }
    async addDataValidation(filename, sheetName, range, validation) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        const validationRule = {
            type: validation.type,
            allowBlank: validation.allowBlank !== false,
            showErrorMessage: validation.showErrorMessage !== false,
            showInputMessage: validation.showInputMessage !== false,
        };
        if (validation.type === 'list' && validation.values) {
            validationRule.formulae = [`"${validation.values.join(',')}"`];
        }
        else if (validation.formula) {
            validationRule.formulae = [validation.formula];
        }
        else if (validation.operator) {
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
        return await workbook.xlsx.writeBuffer();
    }
    async freezePanes(filename, sheetName, row, column) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        if (row && column) {
            worksheet.views = [
                { state: 'frozen', xSplit: column, ySplit: row },
            ];
        }
        else if (row) {
            worksheet.views = [
                { state: 'frozen', ySplit: row },
            ];
        }
        else if (column) {
            worksheet.views = [
                { state: 'frozen', xSplit: column },
            ];
        }
        return await workbook.xlsx.writeBuffer();
    }
    async filterSort(filename, sheetName, range, sortBy, autoFilter = true) {
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
            }
            else {
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
        return await workbook.xlsx.writeBuffer();
    }
    async formatCells(filename, sheetName, range, style) {
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
                    if (style.font)
                        cell.font = style.font;
                    if (style.fill)
                        cell.fill = style.fill;
                    if (style.alignment)
                        cell.alignment = style.alignment;
                    if (style.border)
                        cell.border = style.border;
                    if (style.numFmt)
                        cell.numFmt = style.numFmt;
                }
            }
        }
        return await workbook.xlsx.writeBuffer();
    }
    async addNamedRange(filename, name, range, sheetName) {
        const workbook = await this.loadWorkbook(filename);
        workbook.definedNames.add(sheetName ? `${sheetName}!${range}` : range, name);
        return await workbook.xlsx.writeBuffer();
    }
    async protectSheet(filename, sheetName, password, options) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        worksheet.protect(password || '', options || {});
        return await workbook.xlsx.writeBuffer();
    }
    async mergeWorkbooks(files, outputFilename) {
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
            }
            catch (error) {
                console.error(`Error merging file ${file}:`, error);
            }
        }
        return await workbook.xlsx.writeBuffer();
    }
    async findReplace(filename, find, replace, sheetName, matchCase = false, matchEntireCell = false, searchFormulas = false) {
        const workbook = await this.loadWorkbook(filename);
        const sheets = sheetName
            ? [workbook.getWorksheet(sheetName)]
            : workbook.worksheets;
        sheets.forEach((worksheet) => {
            if (!worksheet)
                return;
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
                        }
                        else {
                            const regex = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
                            cell.value = cellValue.replace(regex, replace);
                        }
                    }
                });
            });
        });
        return await workbook.xlsx.writeBuffer();
    }
    async convertToJSON(excelPath, sheetName, header = true) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelPath);
        const worksheet = sheetName
            ? workbook.getWorksheet(sheetName)
            : workbook.worksheets[0];
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName || 'default'}" not found`);
        }
        const rows = [];
        let headers = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1 && header) {
                headers = row.values;
                headers.shift(); // Remove index 0
            }
            else {
                const values = row.values;
                values.shift(); // Remove index 0
                if (header && headers.length > 0) {
                    const obj = {};
                    headers.forEach((header, idx) => {
                        obj[header] = values[idx];
                    });
                    rows.push(obj);
                }
                else {
                    rows.push(values);
                }
            }
        });
        return JSON.stringify(rows, null, 2);
    }
    async convertToCSV(excelPath, sheetName) {
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
            const values = row.values;
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
    async loadWorkbook(filename) {
        const workbook = new ExcelJS.Workbook();
        try {
            await workbook.xlsx.readFile(filename);
        }
        catch (error) {
            // If file doesn't exist, return empty workbook
            console.warn(`File ${filename} not found, creating new workbook`);
        }
        return workbook;
    }
    applyRowStyle(row, style) {
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
    normalizeColor(color) {
        // Ensure color is in ARGB format
        if (color.startsWith('#')) {
            return 'FF' + color.slice(1).toUpperCase();
        }
        if (color.length === 6) {
            return 'FF' + color.toUpperCase();
        }
        return color.toUpperCase();
    }
    columnToNumber(column) {
        let result = 0;
        for (let i = 0; i < column.length; i++) {
            result = result * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return result;
    }
}
//# sourceMappingURL=excel-generator.js.map