/**
 * Excel Creator - Create beautiful Excel spreadsheets from JSON or Markdown
 * @module excel-creator
 */

import { ExcelBuilder } from './core/excel-builder';
import { JsonParser } from './core/json-parser';
import { MarkdownParser } from './core/markdown-parser';
import { StyleUtils } from './utils/styling';
import { FormulaUtils } from './utils/formulas';
import * as fs from 'fs';
import * as path from 'path';

// Export all types
export * from './types';

// Export core classes
export { ExcelBuilder, JsonParser, MarkdownParser, StyleUtils, FormulaUtils };

/**
 * Create an Excel file from JSON input
 * @param input - JSON string or object containing sheet data
 * @param outputPath - Path where the Excel file will be saved
 * @param options - Additional options for Excel generation
 * @returns Promise that resolves when the file is created
 */
export async function createExcelFromJson(
  input: string | object,
  outputPath: string,
  options: {
    addFormulas?: boolean;
    applyStyling?: boolean;
    preset?: 'minimal' | 'professional' | 'colorful' | 'dark';
  } = {}
): Promise<void> {
  const sheets = JsonParser.parse(input);
  
  const builder = new ExcelBuilder({
    addFormulas: options.addFormulas !== false,
    applyStyling: options.applyStyling !== false
  });

  // Apply preset if specified
  if (options.preset && options.applyStyling !== false) {
    const presetStyles = StyleUtils.getStylePreset(options.preset);
    sheets.forEach(sheet => {
      if (!sheet.config) sheet.config = { name: sheet.name };
      sheet.config.headerStyle = presetStyles.headerStyle;
      sheet.config.dataStyle = presetStyles.dataStyle;
      sheet.config.alternateRowStyle = presetStyles.alternateRowStyle;
    });
  }

  sheets.forEach(sheet => builder.addSheet(sheet));
  await builder.save(outputPath);
}

/**
 * Create an Excel file from Markdown content
 * @param markdown - Markdown string containing tables
 * @param outputPath - Path where the Excel file will be saved
 * @param options - Additional options for Excel generation
 * @returns Promise that resolves when the file is created
 */
export async function createExcelFromMarkdown(
  markdown: string,
  outputPath: string,
  options: {
    addFormulas?: boolean;
    applyStyling?: boolean;
    preset?: 'minimal' | 'professional' | 'colorful' | 'dark';
  } = {}
): Promise<void> {
  const sheets = MarkdownParser.parse(markdown);
  
  const builder = new ExcelBuilder({
    addFormulas: options.addFormulas !== false,
    applyStyling: options.applyStyling !== false
  });

  // Apply preset if specified
  if (options.preset && options.applyStyling !== false) {
    const presetStyles = StyleUtils.getStylePreset(options.preset);
    sheets.forEach(sheet => {
      if (!sheet.config) sheet.config = { name: sheet.name };
      sheet.config.headerStyle = presetStyles.headerStyle;
      sheet.config.dataStyle = presetStyles.dataStyle;
      sheet.config.alternateRowStyle = presetStyles.alternateRowStyle;
    });
  }

  sheets.forEach(sheet => builder.addSheet(sheet));
  await builder.save(outputPath);
}

/**
 * Create an Excel file from a file path (auto-detects format)
 * @param inputPath - Path to the input file (JSON or Markdown)
 * @param outputPath - Path where the Excel file will be saved
 * @param options - Additional options for Excel generation
 * @returns Promise that resolves when the file is created
 */
export async function createExcelFromFile(
  inputPath: string,
  outputPath: string,
  options: {
    addFormulas?: boolean;
    applyStyling?: boolean;
    preset?: 'minimal' | 'professional' | 'colorful' | 'dark';
    format?: 'json' | 'markdown' | 'auto';
  } = {}
): Promise<void> {
  if (!fs.existsSync(inputPath)) {
    throw new Error(`Input file not found: ${inputPath}`);
  }

  const content = fs.readFileSync(inputPath, 'utf-8');
  let format = options.format || 'auto';

  if (format === 'auto') {
    const ext = path.extname(inputPath).toLowerCase();
    if (ext === '.json') {
      format = 'json';
    } else if (ext === '.md' || ext === '.markdown') {
      format = 'markdown';
    } else {
      // Try to detect from content
      try {
        JSON.parse(content);
        format = 'json';
      } catch {
        format = 'markdown';
      }
    }
  }

  if (format === 'markdown') {
    return createExcelFromMarkdown(content, outputPath, options);
  } else {
    return createExcelFromJson(content, outputPath, options);
  }
}

/**
 * Create an Excel file with custom configuration
 * @param config - Complete Excel configuration object
 * @param outputPath - Path where the Excel file will be saved
 * @returns Promise that resolves when the file is created
 */
export async function createExcel(
  config: {
    sheets: Array<{
      name: string;
      headers: string[];
      rows: any[][];
      config?: any;
    }>;
    options?: {
      addFormulas?: boolean;
      applyStyling?: boolean;
      preset?: 'minimal' | 'professional' | 'colorful' | 'dark';
    };
  },
  outputPath: string
): Promise<void> {
  const builder = new ExcelBuilder(config.options || {});

  // Apply preset if specified
  if (config.options?.preset && config.options?.applyStyling !== false) {
    const presetStyles = StyleUtils.getStylePreset(config.options.preset);
    config.sheets.forEach(sheet => {
      if (!sheet.config) sheet.config = { name: sheet.name };
      sheet.config.headerStyle = presetStyles.headerStyle;
      sheet.config.dataStyle = presetStyles.dataStyle;
      sheet.config.alternateRowStyle = presetStyles.alternateRowStyle;
    });
  }

  config.sheets.forEach(sheet => builder.addSheet(sheet));
  await builder.save(outputPath);
}

/**
 * Create an Excel buffer without saving to file
 * @param input - JSON string, object, or sheets configuration
 * @param options - Additional options for Excel generation
 * @returns Promise that resolves with the Excel file buffer
 */
export async function createExcelBuffer(
  input: string | object | { sheets: any[] },
  options: {
    addFormulas?: boolean;
    applyStyling?: boolean;
    preset?: 'minimal' | 'professional' | 'colorful' | 'dark';
    format?: 'json' | 'markdown';
  } = {}
): Promise<Buffer> {
  let sheets: any[];

  if (typeof input === 'object' && 'sheets' in input) {
    sheets = input.sheets;
  } else if (options.format === 'markdown' && typeof input === 'string') {
    sheets = MarkdownParser.parse(input);
  } else {
    sheets = JsonParser.parse(input);
  }

  const builder = new ExcelBuilder({
    addFormulas: options.addFormulas !== false,
    applyStyling: options.applyStyling !== false
  });

  // Apply preset if specified
  if (options.preset && options.applyStyling !== false) {
    const presetStyles = StyleUtils.getStylePreset(options.preset);
    sheets.forEach(sheet => {
      if (!sheet.config) sheet.config = { name: sheet.name };
      sheet.config.headerStyle = presetStyles.headerStyle;
      sheet.config.dataStyle = presetStyles.dataStyle;
      sheet.config.alternateRowStyle = presetStyles.alternateRowStyle;
    });
  }

  sheets.forEach(sheet => builder.addSheet(sheet));
  return builder.toBuffer();
}

/**
 * Quick helper to create a simple Excel file from an array of objects
 * @param data - Array of objects to convert to Excel
 * @param outputPath - Path where the Excel file will be saved
 * @param sheetName - Optional name for the sheet
 * @returns Promise that resolves when the file is created
 */
export async function quickExcel(
  data: Record<string, any>[],
  outputPath: string,
  sheetName: string = 'Data'
): Promise<void> {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Data must be a non-empty array of objects');
  }

  const headers = Object.keys(data[0]);
  const rows = data.map(item => headers.map(header => item[header] ?? ''));

  const builder = new ExcelBuilder({
    addFormulas: true,
    applyStyling: true
  });

  builder.addSheet({
    name: sheetName,
    headers,
    rows,
    config: {
      name: sheetName,
      freezeRows: 1,
      autoFilter: true,
      alternateRows: true
    }
  });

  await builder.save(outputPath);
}

// Default export for convenience
export default {
  createExcelFromJson,
  createExcelFromMarkdown,
  createExcelFromFile,
  createExcel,
  createExcelBuffer,
  quickExcel,
  ExcelBuilder,
  JsonParser,
  MarkdownParser,
  StyleUtils,
  FormulaUtils
};