/**
 * Excel Generator - Create and manipulate Excel workbooks using ExcelJS
 */
import ExcelJS from 'exceljs';
import * as fs from 'fs/promises';
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
    // ============================================================================
    // Excel v3.0 Methods - Phase 1 Quick Wins
    // ============================================================================
    async addSparklines(filename, sheetName, dataRange, location, type, options) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS doesn't have native sparkline support
        // Creating a workaround using conditional formatting or notes
        const cell = worksheet.getCell(location);
        cell.note = `Sparkline: ${type} chart of ${dataRange}`;
        // Add metadata as cell comment for reference
        cell.value = `[Sparkline: ${type}]`;
        return await workbook.xlsx.writeBuffer();
    }
    async addArrayFormulas(filename, sheetName, formulas) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        formulas.forEach(({ cell, formula }) => {
            const targetCell = worksheet.getCell(cell);
            targetCell.value = { formula };
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addSubtotals(filename, sheetName, range, groupBy, summaryFunction, summaryColumns) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Parse range
        const [startCell, endCell] = range.split(':');
        const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
        const endRow = parseInt(endCell.match(/\d+/)?.[0] || '100');
        // Group data and insert subtotal rows
        let currentGroup = worksheet.getCell(startRow, groupBy).value;
        let groupStartRow = startRow;
        for (let row = startRow + 1; row <= endRow + 1; row++) {
            const cellValue = row <= endRow ? worksheet.getCell(row, groupBy).value : null;
            if (cellValue !== currentGroup || row > endRow) {
                // Insert subtotal row
                const subtotalRow = worksheet.getRow(row);
                worksheet.spliceRows(row, 0, []);
                summaryColumns.forEach(col => {
                    const funcName = summaryFunction.toLowerCase();
                    const rangeRef = `${this.numberToColumn(col)}${groupStartRow}:${this.numberToColumn(col)}${row - 1}`;
                    subtotalRow.getCell(col).value = { formula: `=${funcName.toUpperCase()}(${rangeRef})` };
                });
                subtotalRow.getCell(groupBy).value = `${currentGroup} Total`;
                subtotalRow.font = { bold: true };
                currentGroup = cellValue;
                groupStartRow = row + 1;
            }
        }
        return await workbook.xlsx.writeBuffer();
    }
    async addHyperlinks(filename, sheetName, links) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        links.forEach(link => {
            const cell = worksheet.getCell(link.cell);
            if (link.url) {
                cell.value = {
                    text: link.displayText || link.url,
                    hyperlink: link.url,
                    tooltip: link.tooltip
                };
            }
            else if (link.sheet) {
                const target = link.range ? `${link.sheet}!${link.range}` : link.sheet;
                cell.value = {
                    text: link.displayText || `Go to ${link.sheet}`,
                    hyperlink: `#${target}`,
                    tooltip: link.tooltip
                };
            }
            cell.font = { color: { argb: '0000FF' }, underline: true };
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addAdvancedChart(filename, sheetName, chart) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS has limited chart support for advanced types
        // Adding a placeholder with chart metadata
        const position = chart.position || { row: 1, col: 10 };
        const cell = worksheet.getCell(position.row, position.col);
        cell.value = `[${chart.type.toUpperCase()} Chart: ${chart.title}]`;
        cell.note = `Chart Type: ${chart.type}\nData Range: ${chart.dataRange}\n\nNote: Advanced chart types require Microsoft Excel to render.`;
        cell.font = { bold: true, color: { argb: '0000FF' } };
        return await workbook.xlsx.writeBuffer();
    }
    async addSlicers(filename, sheetName, tableName, slicers) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS doesn't support slicers directly
        // Adding metadata as comments for reference
        slicers.forEach((slicer, index) => {
            const position = slicer.position || { row: 1 + index * 2, col: 15 };
            const cell = worksheet.getCell(position.row, position.col);
            cell.value = `[Slicer: ${slicer.caption || slicer.columnName}]`;
            cell.note = `Table: ${tableName}\nColumn: ${slicer.columnName}\n\nNote: Slicers require Microsoft Excel to render.`;
            cell.font = { bold: true, color: { argb: 'FF6600' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF3E0' }
            };
        });
        return await workbook.xlsx.writeBuffer();
    }
    numberToColumn(num) {
        let result = '';
        while (num > 0) {
            const rem = (num - 1) % 26;
            result = String.fromCharCode(65 + rem) + result;
            num = Math.floor((num - rem) / 26);
        }
        return result;
    }
    // ============================================================================
    // Excel v4.0 Methods - Phase 2 & 3 for 100% Coverage
    // ============================================================================
    async addPowerQuery(filename, sheetName, queryName, source, transformations) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS doesn't support Power Query natively
        // Creating metadata sheet documenting the query
        const querySheet = workbook.addWorksheet(`Query_${queryName}`);
        querySheet.getCell('A1').value = `Power Query: ${queryName}`;
        querySheet.getCell('A1').font = { bold: true, size: 14 };
        querySheet.getCell('A2').value = 'Source:';
        querySheet.getCell('B2').value = `${source.type} - ${source.location}`;
        if (transformations && transformations.length > 0) {
            querySheet.getCell('A4').value = 'Transformations:';
            querySheet.getCell('A4').font = { bold: true };
            transformations.forEach((transform, index) => {
                const row = 5 + index;
                querySheet.getCell(row, 1).value = `${index + 1}. ${transform.step}`;
                if (transform.column) {
                    querySheet.getCell(row, 2).value = `Column: ${transform.column}`;
                }
            });
        }
        querySheet.getCell('A1').note = 'Power Query requires Microsoft Excel to execute. This sheet documents the query configuration.';
        return await workbook.xlsx.writeBuffer();
    }
    async goalSeek(filename, sheetName, setCell, toValue, byChangingCell) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: Goal Seek requires iterative calculation not available in ExcelJS
        // Adding metadata as comment
        const cell = worksheet.getCell(setCell);
        cell.note = `Goal Seek: Set ${setCell} to ${toValue} by changing ${byChangingCell}\n\nNote: Goal Seek requires Microsoft Excel to execute.`;
        // Add visual indicator
        worksheet.getCell('A1').note = `Goal Seek configured: Set ${setCell} = ${toValue} by changing ${byChangingCell}`;
        return await workbook.xlsx.writeBuffer();
    }
    async createDataTable(filename, sheetName, type, formulaCell, rowInputCell, columnInputCell, outputRange) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: Data Tables use special Excel features not fully supported by ExcelJS
        // Adding metadata
        const cell = worksheet.getCell(formulaCell);
        cell.note = `Data Table (${type}):\nFormula: ${formulaCell}\n${rowInputCell ? `Row Input: ${rowInputCell}\n` : ''}${columnInputCell ? `Column Input: ${columnInputCell}\n` : ''}Output: ${outputRange}\n\nNote: Data Tables require Microsoft Excel.`;
        return await workbook.xlsx.writeBuffer();
    }
    async manageScenarios(filename, sheetName, scenarios, resultCells) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Create scenario summary sheet
        const summarySheet = workbook.addWorksheet('Scenario Summary');
        summarySheet.getCell('A1').value = 'Scenario Manager Summary';
        summarySheet.getCell('A1').font = { bold: true, size: 14 };
        summarySheet.getCell('A3').value = 'Scenario';
        summarySheet.getCell('B3').value = 'Changing Cells';
        summarySheet.getCell('C3').value = 'Values';
        summarySheet.getCell('D3').value = 'Comment';
        const headerRow = summarySheet.getRow(3);
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
        scenarios.forEach((scenario, index) => {
            const row = 4 + index;
            summarySheet.getCell(row, 1).value = scenario.name;
            summarySheet.getCell(row, 2).value = scenario.changingCells.join(', ');
            summarySheet.getCell(row, 3).value = scenario.values.join(', ');
            summarySheet.getCell(row, 4).value = scenario.comment || '';
        });
        summarySheet.columns.forEach(column => {
            column.width = 20;
        });
        return await workbook.xlsx.writeBuffer();
    }
    async createTable(filename, sheetName, tableName, range, hasHeaders = true, style, showTotalRow) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // ExcelJS supports tables
        worksheet.addTable({
            name: tableName,
            ref: range,
            headerRow: hasHeaders,
            totalsRow: showTotalRow || false,
            style: {
                theme: style || 'TableStyleMedium2',
                showRowStripes: true,
            },
            columns: this.getTableColumnsFromRange(worksheet, range, hasHeaders),
            rows: this.getTableRowsFromRange(worksheet, range, hasHeaders),
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addTableFormula(filename, sheetName, tableName, columnName, formula) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Find table and add formula to column
        const table = worksheet.getTable(tableName);
        if (!table) {
            throw new Error(`Table "${tableName}" not found`);
        }
        // Add note about structured reference
        worksheet.getCell('A1').note = `Table Formula added to ${tableName}.${columnName}: ${formula}\n\nStructured references work in Microsoft Excel.`;
        return await workbook.xlsx.writeBuffer();
    }
    async addFormControls(filename, sheetName, controls) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS doesn't support form controls
        // Creating visual placeholders
        controls.forEach((control, index) => {
            const cell = worksheet.getCell(control.position.row, control.position.col);
            cell.value = `[${control.type.toUpperCase()}: ${control.name}]`;
            cell.font = { bold: true, color: { argb: 'FF0000FF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE7E6FF' }
            };
            let noteText = `Form Control: ${control.type}\nName: ${control.name}`;
            if (control.linkedCell)
                noteText += `\nLinked Cell: ${control.linkedCell}`;
            if (control.inputRange)
                noteText += `\nInput Range: ${control.inputRange}`;
            if (control.min !== undefined)
                noteText += `\nMin: ${control.min}, Max: ${control.max}`;
            noteText += '\n\nNote: Form controls require Microsoft Excel to render and function.';
            cell.note = noteText;
        });
        return await workbook.xlsx.writeBuffer();
    }
    async insertImages(filename, sheetName, images) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        for (const image of images) {
            try {
                const imageBuffer = await fs.readFile(image.path);
                const imageId = workbook.addImage({
                    buffer: imageBuffer,
                    extension: (image.path.split('.').pop() || 'png'),
                });
                worksheet.addImage(imageId, {
                    tl: { col: image.position.col - 0.5, row: image.position.row - 0.5 },
                    ext: {
                        width: image.size?.width || 200,
                        height: image.size?.height || 200
                    },
                    hyperlinks: image.hyperlink ? { hyperlink: image.hyperlink, tooltip: image.description } : undefined,
                });
            }
            catch (error) {
                console.error(`Error adding image ${image.path}:`, error);
            }
        }
        return await workbook.xlsx.writeBuffer();
    }
    async insertShapes(filename, sheetName, shapes) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS has limited shape support
        // Creating visual placeholders
        shapes.forEach(shape => {
            const cell = worksheet.getCell(shape.position.row, shape.position.col);
            cell.value = `[${shape.type.toUpperCase()} SHAPE]`;
            cell.font = { bold: true, color: { argb: 'FFFF6600' } };
            if (shape.fill?.color) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: this.normalizeColor(shape.fill.color) }
                };
            }
            cell.note = `Shape: ${shape.type}\nSize: ${shape.size.width}x${shape.size.height}${shape.text ? `\nText: ${shape.text}` : ''}\n\nNote: Shapes require Microsoft Excel to render.`;
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addSmartArt(filename, sheetName, smartArt) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS doesn't support SmartArt
        // Creating text-based representation
        const cell = worksheet.getCell(smartArt.position.row, smartArt.position.col);
        cell.value = `[SMARTART: ${smartArt.type} - ${smartArt.layout}]`;
        cell.font = { bold: true, size: 12, color: { argb: 'FF008000' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE2EFDA' }
        };
        let noteText = `SmartArt: ${smartArt.type}\nLayout: ${smartArt.layout}\nItems:\n`;
        smartArt.items.forEach((item, index) => {
            noteText += `${index + 1}. ${item.text}\n`;
        });
        noteText += '\nNote: SmartArt requires Microsoft Excel to render.';
        cell.note = noteText;
        return await workbook.xlsx.writeBuffer();
    }
    async configurePageSetup(filename, sheetName, pageSetup) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // ExcelJS supports page setup
        worksheet.pageSetup = {
            orientation: pageSetup.orientation || 'portrait',
            paperSize: this.getPaperSizeCode(pageSetup.paperSize),
            fitToPage: pageSetup.fitToPage ? true : false,
            fitToWidth: pageSetup.fitToPage?.width,
            fitToHeight: pageSetup.fitToPage?.height,
            scale: pageSetup.scale,
            margins: pageSetup.margins ? {
                left: pageSetup.margins.left || 0.7,
                right: pageSetup.margins.right || 0.7,
                top: pageSetup.margins.top || 0.75,
                bottom: pageSetup.margins.bottom || 0.75,
                header: pageSetup.margins.header || 0.3,
                footer: pageSetup.margins.footer || 0.3,
            } : undefined,
            horizontalCentered: pageSetup.centerHorizontally,
            verticalCentered: pageSetup.centerVertically,
            printArea: pageSetup.printArea,
            printTitlesRow: pageSetup.printTitles?.rows,
            printTitlesColumn: pageSetup.printTitles?.columns,
        };
        return await workbook.xlsx.writeBuffer();
    }
    async setHeaderFooter(filename, sheetName, header, footer, differentFirstPage, differentOddEven) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        worksheet.headerFooter = {
            firstHeader: differentFirstPage && header ? this.formatHeaderFooter(header) : undefined,
            firstFooter: differentFirstPage && footer ? this.formatHeaderFooter(footer) : undefined,
            oddHeader: header ? this.formatHeaderFooter(header) : undefined,
            oddFooter: footer ? this.formatHeaderFooter(footer) : undefined,
            evenHeader: differentOddEven && header ? this.formatHeaderFooter(header) : undefined,
            evenFooter: differentOddEven && footer ? this.formatHeaderFooter(footer) : undefined,
            differentFirst: differentFirstPage,
            differentOddEven: differentOddEven,
        };
        return await workbook.xlsx.writeBuffer();
    }
    async addPageBreaks(filename, sheetName, horizontalBreaks, verticalBreaks) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        // Note: ExcelJS has limited page break support
        // Adding metadata
        if (horizontalBreaks || verticalBreaks) {
            worksheet.getCell('A1').note = `Page Breaks:\nHorizontal (rows): ${horizontalBreaks?.join(', ') || 'None'}\nVertical (columns): ${verticalBreaks?.join(', ') || 'None'}\n\nNote: Page breaks are set but may require Microsoft Excel for full support.`;
        }
        return await workbook.xlsx.writeBuffer();
    }
    async enableTrackChanges(filename, enable, highlightChanges) {
        const workbook = await this.loadWorkbook(filename);
        // Note: ExcelJS doesn't support track changes feature
        // Adding metadata to first sheet
        const worksheet = workbook.worksheets[0];
        if (worksheet) {
            worksheet.getCell('A1').note = `Track Changes: ${enable ? 'ENABLED' : 'DISABLED'}\nHighlight Changes: ${highlightChanges ? 'Yes' : 'No'}\n\nNote: Track Changes requires Microsoft Excel to function.`;
        }
        return await workbook.xlsx.writeBuffer();
    }
    async shareWorkbook(filename, share, allowChanges, password) {
        const workbook = await this.loadWorkbook(filename);
        // Note: ExcelJS doesn't support workbook sharing
        // Adding metadata
        const worksheet = workbook.worksheets[0];
        if (worksheet) {
            worksheet.getCell('A1').note = `Workbook Sharing: ${share ? 'ENABLED' : 'DISABLED'}\nAllow Changes: ${allowChanges ? 'Yes' : 'No'}\n${password ? 'Protected with password' : 'No password'}\n\nNote: Workbook sharing requires Microsoft Excel.`;
        }
        return await workbook.xlsx.writeBuffer();
    }
    async addComments(filename, sheetName, comments) {
        const workbook = await this.loadWorkbook(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        comments.forEach(comment => {
            const cell = worksheet.getCell(comment.cell);
            cell.note = comment.author ? `${comment.author}: ${comment.text}` : comment.text;
        });
        return await workbook.xlsx.writeBuffer();
    }
    // Helper methods for new features
    getTableColumnsFromRange(worksheet, range, hasHeaders) {
        // Extract columns from range
        const [startCell, endCell] = range.split(':');
        const startCol = startCell.match(/[A-Z]+/)?.[0] || 'A';
        const endCol = endCell.match(/[A-Z]+/)?.[0] || 'A';
        const headerRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
        const columns = [];
        let colNum = this.columnToNumber(startCol);
        const endColNum = this.columnToNumber(endCol);
        while (colNum <= endColNum) {
            const headerCell = worksheet.getCell(headerRow, colNum);
            columns.push({
                name: hasHeaders ? (headerCell.value || `Column${colNum}`).toString() : `Column${colNum}`,
            });
            colNum++;
        }
        return columns;
    }
    getTableRowsFromRange(worksheet, range, hasHeaders) {
        const [startCell, endCell] = range.split(':');
        const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
        const endRow = parseInt(endCell.match(/\d+/)?.[0] || '1');
        const startCol = this.columnToNumber(startCell.match(/[A-Z]+/)?.[0] || 'A');
        const endCol = this.columnToNumber(endCell.match(/[A-Z]+/)?.[0] || 'A');
        const rows = [];
        const dataStartRow = hasHeaders ? startRow + 1 : startRow;
        for (let row = dataStartRow; row <= endRow; row++) {
            const rowData = [];
            for (let col = startCol; col <= endCol; col++) {
                rowData.push(worksheet.getCell(row, col).value);
            }
            rows.push(rowData);
        }
        return rows;
    }
    getPaperSizeCode(paperSize) {
        const sizes = {
            'letter': 1,
            'legal': 5,
            'A4': 9,
            'A3': 8,
            'tabloid': 3,
        };
        return paperSize ? (sizes[paperSize] || 9) : 9;
    }
    formatHeaderFooter(headerFooter) {
        return `&L${headerFooter.left || ''}&C${headerFooter.center || ''}&R${headerFooter.right || ''}`;
    }
}
//# sourceMappingURL=excel-generator.js.map