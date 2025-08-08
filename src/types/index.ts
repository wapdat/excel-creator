export interface CellStyle {
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: { argb: string };
  };
  fill?: {
    type?: 'pattern';
    pattern?: 'solid' | 'darkVertical' | 'darkHorizontal';
    fgColor?: { argb: string };
    bgColor?: { argb: string };
  };
  border?: {
    top?: BorderStyle;
    left?: BorderStyle;
    bottom?: BorderStyle;
    right?: BorderStyle;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'justify';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  numFmt?: string;
}

export interface BorderStyle {
  style?: 'thin' | 'medium' | 'thick' | 'double';
  color?: { argb: string };
}

export interface ColumnConfig {
  key: string;
  header: string;
  width?: number;
  style?: CellStyle;
  format?: 'currency' | 'percentage' | 'date' | 'number' | 'text';
  formula?: 'sum' | 'average' | 'count' | 'min' | 'max';
}

export interface SheetConfig {
  name: string;
  columns?: ColumnConfig[];
  freezeRows?: number;
  freezeColumns?: number;
  autoFilter?: boolean;
  alternateRows?: boolean;
  headerStyle?: CellStyle;
  dataStyle?: CellStyle;
  alternateRowStyle?: CellStyle;
  showGridLines?: boolean;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
}

export interface ExcelSheet {
  name: string;
  headers: string[];
  rows: any[][];
  config?: SheetConfig;
}

export interface FormulaConfig {
  column: string | number;
  type: 'sum' | 'average' | 'count' | 'min' | 'max';
  startRow?: number;
  endRow?: number;
}

export interface StyleConfig {
  headerStyle?: CellStyle;
  dataStyle?: CellStyle;
  alternateRowStyle?: CellStyle;
  borderStyle?: BorderStyle;
  currencyColumns?: string[];
  percentageColumns?: string[];
  dateColumns?: string[];
}

export interface ExcelConfig {
  sheets: ExcelSheet[];
  globalStyle?: StyleConfig;
  author?: string;
  company?: string;
  created?: Date;
  modified?: Date;
  lastModifiedBy?: string;
  keywords?: string[];
  category?: string;
  description?: string;
}

export interface JsonInput {
  sheets?: ExcelSheet[];
  data?: any[][] | Record<string, any>[];
  headers?: string[];
  config?: SheetConfig | SheetConfig[];
  globalConfig?: ExcelConfig;
}

export interface MarkdownTable {
  headers: string[];
  rows: string[][];
  title?: string;
}

export interface ParseOptions {
  format?: 'json' | 'markdown' | 'auto';
  encoding?: BufferEncoding;
  delimiter?: string;
  quote?: string;
  escape?: string;
}

export interface ExportOptions {
  outputPath: string;
  overwrite?: boolean;
  openAfterExport?: boolean;
  verbose?: boolean;
  dryRun?: boolean;
}

export interface CliOptions {
  input?: string;
  output?: string;
  format?: 'json' | 'markdown' | 'auto';
  config?: string;
  sheet?: string;
  noFormulas?: boolean;
  noStyling?: boolean;
  formulas?: boolean;
  styling?: boolean;
  freezeRows?: number;
  alternateRows?: boolean;
  currency?: string;
  preset?: string;
  debug?: boolean;
  examples?: boolean;
  overwrite?: boolean;
  openAfterExport?: boolean;
  verbose?: boolean;
  dryRun?: boolean;
}

export interface BuilderOptions {
  addFormulas?: boolean;
  applyStyling?: boolean;
  defaultSheetName?: string;
  dateFormat?: string;
  currencySymbol?: string;
  decimalPlaces?: number;
}

export type CellValue = string | number | boolean | Date | null | undefined;

export interface ValidationRule {
  type: 'list' | 'decimal' | 'integer' | 'date' | 'textLength' | 'custom';
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  formulae?: any[];
  operator?: string;
}

export interface ConditionalFormat {
  ref: string;
  rules: Array<{
    type: string;
    priority: number;
    style?: CellStyle;
    formulae?: string[];
  }>;
}