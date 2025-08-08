import { ExcelSheet, JsonInput, SheetConfig, ColumnConfig } from '../types';

export class JsonParser {
  /**
   * Parse JSON input and convert to ExcelSheet format
   */
  static parse(input: string | JsonInput): ExcelSheet[] {
    let data: JsonInput;

    if (typeof input === 'string') {
      try {
        data = JSON.parse(input);
      } catch (error) {
        throw new Error(`Invalid JSON input: ${(error as Error).message}`);
      }
    } else {
      data = input;
    }

    if (data.sheets && Array.isArray(data.sheets)) {
      return data.sheets.map(sheet => this.validateSheet(sheet));
    }

    if (data.data) {
      return [this.createSheetFromData(data)];
    }

    throw new Error('Invalid input: must contain either "sheets" array or "data" field');
  }

  /**
   * Create a sheet from simple data array
   */
  private static createSheetFromData(input: JsonInput): ExcelSheet {
    const { data, headers, config } = input;
    
    if (!data || !Array.isArray(data)) {
      throw new Error('Data must be an array');
    }

    if (data.length === 0) {
      return {
        name: (config as SheetConfig)?.name || 'Sheet1',
        headers: headers || [],
        rows: [],
        config: config as SheetConfig
      };
    }

    let sheetHeaders: string[];
    let sheetRows: any[][];

    if (this.isObjectArray(data)) {
      sheetHeaders = headers || Object.keys(data[0] as Record<string, any>);
      sheetRows = (data as Record<string, any>[]).map(row => 
        sheetHeaders.map(header => row[header] ?? '')
      );
    } else {
      sheetHeaders = headers || this.generateHeaders((data[0] as any[]).length);
      sheetRows = data as any[][];
    }

    return {
      name: (config as SheetConfig)?.name || 'Sheet1',
      headers: sheetHeaders,
      rows: sheetRows,
      config: config as SheetConfig
    };
  }

  /**
   * Validate and normalize a sheet object
   */
  private static validateSheet(sheet: ExcelSheet): ExcelSheet {
    if (!sheet.name) {
      throw new Error('Sheet must have a name');
    }

    if (!Array.isArray(sheet.headers)) {
      throw new Error(`Sheet "${sheet.name}" must have headers array`);
    }

    if (!Array.isArray(sheet.rows)) {
      throw new Error(`Sheet "${sheet.name}" must have rows array`);
    }

    if (sheet.rows.length > 0) {
      const expectedColumns = sheet.headers.length;
      sheet.rows.forEach((row, index) => {
        if (!Array.isArray(row)) {
          throw new Error(`Row ${index} in sheet "${sheet.name}" must be an array`);
        }
        if (row.length !== expectedColumns) {
          console.warn(`Row ${index} in sheet "${sheet.name}" has ${row.length} columns, expected ${expectedColumns}`);
          while (row.length < expectedColumns) {
            row.push('');
          }
          if (row.length > expectedColumns) {
            row.length = expectedColumns;
          }
        }
      });
    }

    return sheet;
  }

  /**
   * Check if data is an array of objects
   */
  private static isObjectArray(data: any[]): boolean {
    return data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0]);
  }

  /**
   * Generate default headers for numeric columns
   */
  private static generateHeaders(count: number): string[] {
    const headers: string[] = [];
    for (let i = 0; i < count; i++) {
      headers.push(this.numberToColumn(i));
    }
    return headers;
  }

  /**
   * Convert number to Excel column letter (0 -> A, 1 -> B, 26 -> AA, etc.)
   */
  private static numberToColumn(num: number): string {
    let column = '';
    let n = num;
    
    while (n >= 0) {
      column = String.fromCharCode(65 + (n % 26)) + column;
      n = Math.floor(n / 26) - 1;
      if (n < 0) break;
    }
    
    return column;
  }

  /**
   * Parse a configuration file
   */
  static parseConfig(configPath: string): SheetConfig | SheetConfig[] {
    try {
      const fs = require('fs');
      const configContent = fs.readFileSync(configPath, 'utf-8');
      return JSON.parse(configContent);
    } catch (error) {
      throw new Error(`Failed to parse config file: ${(error as Error).message}`);
    }
  }

  /**
   * Merge configurations
   */
  static mergeConfig(base?: SheetConfig, override?: SheetConfig): SheetConfig | undefined {
    if (!base && !override) return undefined;
    if (!base) return override;
    if (!override) return base;

    return {
      ...base,
      ...override,
      headerStyle: { ...base.headerStyle, ...override.headerStyle },
      dataStyle: { ...base.dataStyle, ...override.dataStyle },
      alternateRowStyle: { ...base.alternateRowStyle, ...override.alternateRowStyle },
      columns: override.columns || base.columns
    };
  }

  /**
   * Infer column configurations from data
   */
  static inferColumns(headers: string[], rows: any[][]): ColumnConfig[] {
    const columns: ColumnConfig[] = [];
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const columnData = rows.map(row => row[i]).filter(val => val != null && val !== '');
      
      const config: ColumnConfig = {
        key: header,
        header: header,
        width: Math.max(header.length * 1.2, 10)
      };

      if (columnData.length > 0) {
        const dataType = this.inferDataType(columnData);
        
        switch (dataType) {
          case 'currency':
            config.format = 'currency';
            config.formula = 'sum';
            config.width = Math.max(config.width ?? 10, 15);
            break;
          case 'percentage':
            config.format = 'percentage';
            config.width = Math.max(config.width ?? 10, 12);
            break;
          case 'date':
            config.format = 'date';
            config.width = Math.max(config.width ?? 10, 12);
            break;
          case 'number':
            config.format = 'number';
            config.formula = 'sum';
            config.width = Math.max(config.width ?? 10, 12);
            break;
          default:
            config.format = 'text';
            const maxLength = Math.max(...columnData.map(d => String(d).length));
            config.width = Math.min(Math.max(config.width ?? 10, maxLength * 1.1), 50);
        }
      }

      columns.push(config);
    }

    return columns;
  }

  /**
   * Infer data type from column values
   */
  private static inferDataType(values: any[]): 'currency' | 'percentage' | 'date' | 'number' | 'text' {
    const sample = values.slice(0, Math.min(10, values.length));
    
    const isCurrency = sample.every(val => {
      const str = String(val);
      return /^[$£€¥]?[\d,]+\.?\d*$/.test(str) || /^\d+\.?\d*$/.test(str);
    });
    
    if (isCurrency && sample.some(val => String(val).includes('$') || String(val).includes('£'))) {
      return 'currency';
    }

    const isPercentage = sample.every(val => {
      const str = String(val);
      return /^\d+\.?\d*%$/.test(str) || (typeof val === 'number' && val >= 0 && val <= 1);
    });
    
    if (isPercentage) {
      return 'percentage';
    }

    const isDate = sample.every(val => {
      const date = new Date(val);
      return !isNaN(date.getTime()) && String(val).includes('-') || String(val).includes('/');
    });
    
    if (isDate) {
      return 'date';
    }

    const isNumber = sample.every(val => {
      return typeof val === 'number' || /^-?\d+\.?\d*$/.test(String(val));
    });
    
    if (isNumber) {
      return 'number';
    }

    return 'text';
  }
}