import * as ExcelJS from 'exceljs';
import { 
  ExcelSheet, 
  SheetConfig, 
  CellStyle, 
  BuilderOptions,
  ColumnConfig
} from '../types';
import { StyleUtils } from '../utils/styling';
import { FormulaUtils } from '../utils/formulas';
import { JsonParser } from './json-parser';

export class ExcelBuilder {
  private workbook: ExcelJS.Workbook;
  private options: Required<BuilderOptions>;

  constructor(options: BuilderOptions = {}) {
    this.workbook = new ExcelJS.Workbook();
    this.options = {
      addFormulas: options.addFormulas !== false,
      applyStyling: options.applyStyling !== false,
      defaultSheetName: options.defaultSheetName || 'Sheet',
      dateFormat: options.dateFormat || 'yyyy-mm-dd',
      currencySymbol: options.currencySymbol || '$',
      decimalPlaces: options.decimalPlaces ?? 2
    };

    // Set workbook properties
    this.workbook.creator = 'Excel Creator';
    this.workbook.created = new Date();
    this.workbook.modified = new Date();
  }

  /**
   * Add a sheet to the workbook
   */
  addSheet(sheet: ExcelSheet): void {
    const sheetName = sheet.name || this.options.defaultSheetName;
    const worksheet = this.workbook.addWorksheet(sheetName);
    const config: SheetConfig = sheet.config || { 
      name: sheetName
    };

    // Infer column configurations if not provided
    const columns = config.columns || JsonParser.inferColumns(sheet.headers, sheet.rows);

    // Add headers
    this.addHeaders(worksheet, sheet.headers, config, columns);

    // Add data rows
    this.addDataRows(worksheet, sheet.rows, config, columns);

    // Add formulas if enabled
    if (this.options.addFormulas && sheet.rows.length > 0) {
      this.addFormulas(worksheet, sheet.headers, sheet.rows, columns);
    }

    // Apply sheet configuration
    this.applySheetConfig(worksheet, config, sheet.headers.length, sheet.rows.length);

    // Set column widths
    this.setColumnWidths(worksheet, columns, sheet.headers);

    // Apply conditional formatting if needed
    if (config.alternateRows && this.options.applyStyling) {
      this.applyAlternateRowStyling(worksheet, sheet.rows.length, config);
    }
  }

  /**
   * Add headers to worksheet
   */
  private addHeaders(
    worksheet: ExcelJS.Worksheet,
    headers: string[],
    config: SheetConfig,
    columns: ColumnConfig[]
  ): void {
    const headerRow = worksheet.addRow(headers);

    if (this.options.applyStyling) {
      const headerStyle = config.headerStyle || StyleUtils.getDefaultHeaderStyle();
      
      headerRow.eachCell((cell, colNumber) => {
        const column = columns[colNumber - 1];
        const cellStyle = StyleUtils.mergeStyles(
          headerStyle,
          column?.style
        );
        
        if (cellStyle) {
          this.applyCellStyle(cell, cellStyle);
        }
      });

      headerRow.height = 25;
    }
  }

  /**
   * Add data rows to worksheet
   */
  private addDataRows(
    worksheet: ExcelJS.Worksheet,
    rows: any[][],
    config: SheetConfig,
    columns: ColumnConfig[]
  ): void {
    rows.forEach((rowData) => {
      const row = worksheet.addRow(rowData);

      if (this.options.applyStyling) {
        const dataStyle = config.dataStyle || StyleUtils.getDefaultDataStyle();
        
        row.eachCell((cell, colNumber) => {
          const column = columns[colNumber - 1];
          const value = rowData[colNumber - 1];

          // Apply base data style
          let cellStyle = dataStyle;

          // Apply column-specific formatting
          if (column?.format) {
            const formatStyle = this.getFormatStyle(column.format);
            cellStyle = StyleUtils.mergeStyles(cellStyle, formatStyle) || cellStyle;
          }

          // Apply column-specific style
          if (column?.style) {
            cellStyle = StyleUtils.mergeStyles(cellStyle, column.style) || cellStyle;
          }

          if (cellStyle) {
            this.applyCellStyle(cell, cellStyle);
          }

          // Handle special value types
          this.handleCellValue(cell, value, column?.format);
        });
      }
    });
  }

  /**
   * Get format-specific style
   */
  private getFormatStyle(format: string): CellStyle | undefined {
    switch (format) {
      case 'currency':
        return StyleUtils.getCurrencyStyle(this.options.currencySymbol);
      case 'percentage':
        return StyleUtils.getPercentageStyle(this.options.decimalPlaces);
      case 'date':
        return StyleUtils.getDateStyle(this.options.dateFormat);
      case 'number':
        return StyleUtils.getNumberStyle(this.options.decimalPlaces);
      default:
        return undefined;
    }
  }

  /**
   * Handle special cell value types
   */
  private handleCellValue(cell: ExcelJS.Cell, value: any, format?: string): void {
    if (value == null || value === '') {
      return;
    }

    // Handle dates
    if (format === 'date' || value instanceof Date) {
      if (!(value instanceof Date)) {
        value = new Date(value);
      }
      if (!isNaN(value.getTime())) {
        cell.value = value;
      }
    }
    // Handle percentages
    else if (format === 'percentage') {
      if (typeof value === 'string' && value.endsWith('%')) {
        cell.value = parseFloat(value) / 100;
      } else if (typeof value === 'number') {
        cell.value = value <= 1 ? value : value / 100;
      }
    }
    // Handle currency
    else if (format === 'currency') {
      const cleaned = String(value).replace(/[$£€¥,]/g, '').trim();
      const numValue = parseFloat(cleaned);
      if (!isNaN(numValue)) {
        cell.value = numValue;
      }
    }
    // Handle numbers
    else if (format === 'number' || typeof value === 'number') {
      const numValue = typeof value === 'number' ? value : parseFloat(value);
      if (!isNaN(numValue)) {
        cell.value = numValue;
      }
    }
  }

  /**
   * Add formulas to the worksheet
   */
  private addFormulas(
    worksheet: ExcelJS.Worksheet,
    _headers: string[],
    rows: any[][],
    columns: ColumnConfig[]
  ): void {
    if (rows.length === 0) return;

    const formulaRow = worksheet.addRow([]);
    let hasFormulas = false;

    columns.forEach((column, colIndex) => {
      const columnLetter = FormulaUtils.numberToColumn(colIndex + 1);
      const startCell = `${columnLetter}2`;
      const endCell = `${columnLetter}${rows.length + 1}`;
      
      // Check if column contains numeric data
      const columnData = rows.map(row => row[colIndex]);
      const isNumeric = FormulaUtils.isNumericColumn(columnData);

      if (column.formula && isNumeric) {
        const formula = FormulaUtils.generateFormula(
          column.formula,
          startCell,
          endCell
        );
        formulaRow.getCell(colIndex + 1).value = { formula: `=${formula}` };
        hasFormulas = true;
      } else if (column.formula === 'sum' && isNumeric) {
        // Default to SUM for numeric columns without specific formula
        const formula = FormulaUtils.sum(startCell, endCell);
        formulaRow.getCell(colIndex + 1).value = { formula: `=${formula}` };
        hasFormulas = true;
      } else {
        formulaRow.getCell(colIndex + 1).value = '';
      }
    });

    // Apply total row styling if formulas were added
    if (hasFormulas && this.options.applyStyling) {
      const totalStyle = StyleUtils.getTotalRowStyle();
      formulaRow.eachCell((cell) => {
        if (cell.value) {
          this.applyCellStyle(cell, totalStyle);
        }
      });
      formulaRow.height = 22;
    }

    // If no formulas were added, remove the empty row
    if (!hasFormulas) {
      worksheet.spliceRows(formulaRow.number, 1);
    }
  }

  /**
   * Apply cell style
   */
  private applyCellStyle(cell: ExcelJS.Cell, style: CellStyle): void {
    if (style.font) {
      cell.font = style.font as ExcelJS.Font;
    }
    if (style.fill) {
      cell.fill = style.fill as ExcelJS.Fill;
    }
    if (style.border) {
      cell.border = style.border as ExcelJS.Borders;
    }
    if (style.alignment) {
      cell.alignment = style.alignment as ExcelJS.Alignment;
    }
    if (style.numFmt) {
      cell.numFmt = style.numFmt;
    }
  }

  /**
   * Apply sheet configuration
   */
  private applySheetConfig(
    worksheet: ExcelJS.Worksheet,
    config: SheetConfig,
    columnCount: number,
    rowCount: number
  ): void {
    // Freeze rows/columns
    if (config.freezeRows || config.freezeColumns) {
      worksheet.views = [{
        state: 'frozen',
        xSplit: config.freezeColumns || 0,
        ySplit: config.freezeRows || 0,
        activeCell: 'A1'
      }];
    }

    // Auto filter
    if (config.autoFilter && rowCount > 0) {
      const lastColumn = FormulaUtils.numberToColumn(columnCount);
      worksheet.autoFilter = {
        from: 'A1',
        to: `${lastColumn}${rowCount + 1}`
      };
    }

    // Grid lines
    if (config.showGridLines !== undefined) {
      worksheet.views = worksheet.views || [{}];
      worksheet.views[0].showGridLines = config.showGridLines;
    }

    // Default row height
    if (config.defaultRowHeight) {
      worksheet.properties.defaultRowHeight = config.defaultRowHeight;
    }
  }

  /**
   * Set column widths
   */
  private setColumnWidths(
    worksheet: ExcelJS.Worksheet,
    columns: ColumnConfig[],
    headers: string[]
  ): void {
    columns.forEach((column, index) => {
      const worksheetColumn = worksheet.getColumn(index + 1);
      
      if (column.width) {
        worksheetColumn.width = column.width;
      } else {
        // Auto-size based on header length
        const headerLength = headers[index]?.length || 10;
        worksheetColumn.width = Math.max(headerLength * 1.2, 10);
      }
    });
  }

  /**
   * Apply alternate row styling
   */
  private applyAlternateRowStyling(
    worksheet: ExcelJS.Worksheet,
    rowCount: number,
    config: SheetConfig
  ): void {
    const alternateStyle = config.alternateRowStyle || StyleUtils.getAlternateRowStyle();
    
    for (let i = 3; i <= rowCount + 1; i += 2) {
      const row = worksheet.getRow(i);
      row.eachCell((cell) => {
        const currentStyle: CellStyle = {
          font: cell.font as any,
          border: cell.border as any,
          alignment: cell.alignment as any,
          numFmt: cell.numFmt as string
        };
        const mergedStyle = StyleUtils.mergeStyles(currentStyle, alternateStyle);
        if (mergedStyle) {
          this.applyCellStyle(cell, mergedStyle);
        }
      });
    }
  }

  /**
   * Save the workbook to a file
   */
  async save(filePath: string): Promise<void> {
    await this.workbook.xlsx.writeFile(filePath);
  }

  /**
   * Get the workbook as a buffer
   */
  async toBuffer(): Promise<Buffer> {
    const buffer = await this.workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  /**
   * Get the underlying ExcelJS workbook
   */
  getWorkbook(): ExcelJS.Workbook {
    return this.workbook;
  }

  /**
   * Create Excel from JSON input
   */
  static async fromJson(
    input: string | object,
    options: BuilderOptions = {}
  ): Promise<ExcelBuilder> {
    const sheets = JsonParser.parse(input);
    const builder = new ExcelBuilder(options);
    
    sheets.forEach(sheet => {
      builder.addSheet(sheet);
    });
    
    return builder;
  }

  /**
   * Set workbook properties
   */
  setProperties(properties: Partial<{
    creator: string;
    lastModifiedBy: string;
    created: Date;
    modified: Date;
    company: string;
    title: string;
    subject: string;
    keywords: string;
    category: string;
    description: string;
    manager: string;
  }>): void {
    Object.assign(this.workbook, properties);
  }
}