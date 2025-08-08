#!/usr/bin/env node

import { Command } from 'commander';
import * as fs from 'fs';
import * as path from 'path';
import chalk from 'chalk';
import ora from 'ora';
import { version } from '../package.json';
import { ExcelBuilder } from './core/excel-builder';
import { JsonParser } from './core/json-parser';
import { MarkdownParser } from './core/markdown-parser';
import { CliHelp } from './utils/cli-help';
import { StyleUtils } from './utils/styling';
import { CliOptions, ExcelSheet, SheetConfig } from './types';

const program = new Command();

program
  .name('excel-creator')
  .description('Create beautiful Excel spreadsheets from JSON or Markdown')
  .version(version, '-v, --version', 'Show version number')
  .helpOption('-h, --help', 'Show help')
  .option('-i, --input <path>', 'Input file (JSON/Markdown)')
  .option('-o, --output <path>', 'Output Excel file', 'output.xlsx')
  .option('-f, --format <type>', 'Input format: json|markdown|auto', 'auto')
  .option('-c, --config <path>', 'Configuration file for styling')
  .option('--sheet <name>', 'Sheet name for single-sheet files')
  .option('--no-formulas', 'Disable automatic SUM formulas')
  .option('--no-styling', 'Disable default styling')
  .option('--freeze-rows <n>', 'Number of rows to freeze', parseInt)
  .option('--alternate-rows', 'Enable alternating row colors')
  .option('--currency <columns>', 'Comma-separated column names for currency format')
  .option('--preset <style>', 'Style preset: minimal|professional|colorful|dark', 'professional')
  .option('--examples', 'Show usage examples')
  .option('--debug', 'Enable debug output')
  .action(async (options: CliOptions) => {
    // Show examples if requested
    if (options.examples) {
      CliHelp.showBanner();
      CliHelp.showExamples();
      CliHelp.showJsonFormat();
      CliHelp.showMarkdownFormat();
      CliHelp.showConfigFormat();
      CliHelp.showFooter();
      process.exit(0);
    }

    // Show banner for interactive mode
    if (!options.input && process.stdin.isTTY) {
      CliHelp.showBanner();
      CliHelp.showUsage();
      CliHelp.showOptions();
      CliHelp.showFooter();
      process.exit(0);
    }

    try {
      await processInput(options);
    } catch (error) {
      handleError(error as Error, options);
    }
  });

/**
 * Process input and create Excel file
 */
async function processInput(options: CliOptions): Promise<void> {
  const spinner = ora(CliHelp.getSpinnerText('Processing input')).start();

  try {
    // Read input
    let inputContent: string;
    if (options.input) {
      if (!fs.existsSync(options.input)) {
        throw new Error(`Input file not found: ${options.input}`);
      }
      inputContent = fs.readFileSync(options.input, 'utf-8');
      spinner.text = CliHelp.getSpinnerText(`Reading ${options.input}`);
    } else {
      inputContent = await readStdin();
      spinner.text = CliHelp.getSpinnerText('Reading from stdin');
    }

    if (options.debug) {
      spinner.stop();
      CliHelp.showDebug('Input length', inputContent.length);
      spinner.start();
    }

    // Detect format
    let format = options.format;
    if (format === 'auto') {
      format = detectFormat(inputContent, options.input);
      spinner.text = CliHelp.getSpinnerText(`Detected format: ${format}`);
    }

    // Parse input
    let sheets: ExcelSheet[];
    if (format === 'markdown') {
      sheets = MarkdownParser.parse(inputContent);
    } else {
      sheets = JsonParser.parse(inputContent);
    }

    spinner.text = CliHelp.getSpinnerText(`Parsed ${sheets.length} sheet(s)`);

    // Apply configuration
    if (options.config) {
      const config = JsonParser.parseConfig(options.config);
      applyConfiguration(sheets, config);
    }

    // Apply CLI options
    applyCliOptions(sheets, options);

    // Create Excel file
    spinner.text = CliHelp.getSpinnerText('Creating Excel file');
    const builder = new ExcelBuilder({
      addFormulas: !options.noFormulas,
      applyStyling: !options.noStyling
    });

    // Apply preset styling
    if (!options.noStyling && options.preset) {
      const presetStyles = StyleUtils.getStylePreset(options.preset as any);
      sheets.forEach(sheet => {
        if (!sheet.config) sheet.config = { name: sheet.name };
        sheet.config.headerStyle = presetStyles.headerStyle;
        sheet.config.dataStyle = presetStyles.dataStyle;
        sheet.config.alternateRowStyle = presetStyles.alternateRowStyle;
      });
    }

    // Add sheets to workbook
    sheets.forEach(sheet => {
      builder.addSheet(sheet);
    });

    // Save file
    const outputPath = options.output || 'output.xlsx';
    spinner.text = CliHelp.getSpinnerText(`Saving to ${outputPath}`);
    await builder.save(outputPath);

    // Calculate statistics
    const totalRows = sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);

    spinner.succeed(chalk.green('Excel file created successfully!'));
    CliHelp.showSuccess(outputPath, {
      sheets: sheets.length,
      rows: totalRows
    });

  } catch (error) {
    spinner.fail(chalk.red('Failed to create Excel file'));
    throw error;
  }
}

/**
 * Read from stdin
 */
function readStdin(): Promise<string> {
  return new Promise((resolve, reject) => {
    let data = '';
    
    process.stdin.setEncoding('utf-8');
    
    process.stdin.on('data', chunk => {
      data += chunk;
    });
    
    process.stdin.on('end', () => {
      resolve(data);
    });
    
    process.stdin.on('error', err => {
      reject(err);
    });

    // Set timeout for stdin
    setTimeout(() => {
      if (data === '') {
        reject(new Error('No input received from stdin'));
      }
    }, 5000);
  });
}

/**
 * Detect input format
 */
function detectFormat(content: string, filePath?: string): 'json' | 'markdown' {
  // Check file extension
  if (filePath) {
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.json') return 'json';
    if (ext === '.md' || ext === '.markdown') return 'markdown';
  }

  // Try to parse as JSON
  try {
    JSON.parse(content);
    return 'json';
  } catch {
    // Check for markdown table indicators
    if (content.includes('|') && content.includes('-')) {
      return 'markdown';
    }
  }

  // Default to JSON
  return 'json';
}

/**
 * Apply configuration to sheets
 */
function applyConfiguration(sheets: ExcelSheet[], config: SheetConfig | SheetConfig[]): void {
  if (Array.isArray(config)) {
    config.forEach((sheetConfig, index) => {
      if (sheets[index]) {
        sheets[index].config = JsonParser.mergeConfig(sheets[index].config, sheetConfig);
      }
    });
  } else {
    sheets.forEach(sheet => {
      sheet.config = JsonParser.mergeConfig(sheet.config, config);
    });
  }
}

/**
 * Apply CLI options to sheets
 */
function applyCliOptions(sheets: ExcelSheet[], options: CliOptions): void {
  sheets.forEach(sheet => {
    if (!sheet.config) {
      sheet.config = { name: sheet.name };
    }

    // Apply sheet name if single sheet
    if (sheets.length === 1 && options.sheet) {
      sheet.name = options.sheet;
    }

    // Apply freeze rows
    if (options.freezeRows !== undefined) {
      sheet.config.freezeRows = options.freezeRows;
    }

    // Apply alternate rows
    if (options.alternateRows) {
      sheet.config.alternateRows = true;
    }

    // Apply currency columns
    if (options.currency) {
      const currencyColumns = options.currency.split(',').map(c => c.trim());
      if (!sheet.config.columns) {
        sheet.config.columns = [];
      }
      
      currencyColumns.forEach(colName => {
        const colIndex = sheet.headers.indexOf(colName);
        if (colIndex >= 0) {
          const existingCol = sheet.config!.columns!.find(c => c.key === colName);
          if (existingCol) {
            existingCol.format = 'currency';
            existingCol.formula = 'sum';
          } else {
            sheet.config!.columns!.push({
              key: colName,
              header: colName,
              format: 'currency',
              formula: 'sum'
            });
          }
        }
      });
    }
  });
}

/**
 * Handle errors
 */
function handleError(error: Error, options: CliOptions): void {
  let suggestion: string | undefined;

  if (error.message.includes('not found')) {
    suggestion = 'Check that the file path is correct and the file exists.';
  } else if (error.message.includes('Invalid JSON')) {
    suggestion = 'Ensure your JSON is properly formatted. You can validate it at jsonlint.com';
  } else if (error.message.includes('No tables found')) {
    suggestion = 'Make sure your markdown file contains properly formatted tables.';
  } else if (error.message.includes('Permission denied')) {
    suggestion = 'Check that you have write permissions for the output directory.';
  }

  if (options.debug) {
    console.error(chalk.red('\nFull error:'));
    console.error(error);
  }

  CliHelp.showError(error.message, suggestion);
  process.exit(1);
}

// Parse arguments
program.parse(process.argv);