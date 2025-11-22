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
    async addFormulas(filename, sheetName, formulas) {
        const workbook = new ExcelJS.Workbook();
        // In production, you'd load the existing file
        // For now, create a new workbook
        const worksheet = workbook.addWorksheet(sheetName);
        formulas.forEach(({ cell, formula }) => {
            const cellObj = worksheet.getCell(cell);
            cellObj.value = { formula };
        });
        return await workbook.xlsx.writeBuffer();
    }
    async addChart(filename, sheetName, chart) {
        // ExcelJS has limited chart support, so we'll note this in documentation
        // as requiring manual post-processing or using python-excel libraries
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(sheetName);
        // Add a note about chart creation
        worksheet.getCell('A1').value = `Chart "${chart.title}" would be created here`;
        worksheet.getCell('A1').note = `Chart Type: ${chart.type}, Data Range: ${chart.dataRange}`;
        return await workbook.xlsx.writeBuffer();
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
    async convertToCSV(excelPath, sheetName) {
        const workbook = new ExcelJS.Workbook();
        // In production, load from excelPath
        const worksheet = workbook.worksheets[0] || workbook.addWorksheet('Sheet1');
        let csv = '';
        worksheet.eachRow((row, rowNumber) => {
            const values = row.values;
            csv += values.slice(1).join(',') + '\n'; // Skip index 0
        });
        return csv;
    }
}
//# sourceMappingURL=excel-generator.js.map