import { ExcelSheet, MarkdownTable } from '../types';

export class MarkdownParser {
  /**
   * Parse markdown content and extract tables
   */
  static parse(content: string): ExcelSheet[] {
    const tables = this.extractTables(content);
    
    if (tables.length === 0) {
      throw new Error('No tables found in markdown content');
    }

    return tables.map((table, index) => this.tableToSheet(table, index));
  }

  /**
   * Extract all tables from markdown content
   */
  private static extractTables(content: string): MarkdownTable[] {
    const tables: MarkdownTable[] = [];
    const lines = content.split('\n');
    
    let i = 0;
    while (i < lines.length) {
      if (this.isTableStart(lines, i)) {
        const table = this.parseTable(lines, i);
        if (table) {
          tables.push(table);
          i += table.rows.length + 2; // Skip past the parsed table
        }
      }
      i++;
    }

    // Also try to extract HTML tables if any
    const htmlTables = this.extractHtmlTables(content);
    tables.push(...htmlTables);

    return tables;
  }

  /**
   * Check if current line starts a markdown table
   */
  private static isTableStart(lines: string[], index: number): boolean {
    if (index + 1 >= lines.length) return false;
    
    const currentLine = lines[index].trim();
    const nextLine = lines[index + 1].trim();
    
    // Check for pipe-separated values and separator line
    const hasPipes = currentLine.includes('|');
    const isSeparator = /^[\|\s\-:]+$/.test(nextLine) && nextLine.includes('|');
    
    return hasPipes && isSeparator;
  }

  /**
   * Parse a markdown table starting at the given index
   */
  private static parseTable(lines: string[], startIndex: number): MarkdownTable | null {
    const headerLine = lines[startIndex].trim();
    const separatorLine = lines[startIndex + 1].trim();
    
    if (!this.isValidSeparator(separatorLine)) {
      return null;
    }

    // Find title if it exists (line above the table)
    let title: string | undefined;
    if (startIndex > 0) {
      const prevLine = lines[startIndex - 1].trim();
      if (prevLine && !prevLine.includes('|')) {
        // Check if it's a heading
        if (prevLine.startsWith('#')) {
          title = prevLine.replace(/^#+\s*/, '');
        } else if (prevLine.length < 100) {
          title = prevLine;
        }
      }
    }

    // Parse headers
    const headers = this.parseTableRow(headerLine);
    if (headers.length === 0) return null;

    // Parse data rows
    const rows: string[][] = [];
    let i = startIndex + 2;
    
    while (i < lines.length) {
      const line = lines[i].trim();
      if (!line || !line.includes('|')) break;
      
      const row = this.parseTableRow(line);
      if (row.length === headers.length) {
        rows.push(row);
      } else if (row.length > 0) {
        // Pad or trim row to match header length
        while (row.length < headers.length) row.push('');
        if (row.length > headers.length) row.length = headers.length;
        rows.push(row);
      }
      
      i++;
    }

    return { headers, rows, title };
  }

  /**
   * Check if a line is a valid table separator
   */
  private static isValidSeparator(line: string): boolean {
    // Must contain pipes and dashes/colons
    if (!line.includes('|') || !line.includes('-')) return false;
    
    // Remove all valid separator characters and check if anything remains
    const cleaned = line.replace(/[\|\s\-:]/g, '');
    return cleaned.length === 0;
  }

  /**
   * Parse a single table row
   */
  private static parseTableRow(line: string): string[] {
    // Remove leading and trailing pipes if present
    let cleaned = line.trim();
    if (cleaned.startsWith('|')) cleaned = cleaned.substring(1);
    if (cleaned.endsWith('|')) cleaned = cleaned.substring(0, cleaned.length - 1);
    
    // Split by pipe and clean each cell
    return cleaned
      .split('|')
      .map(cell => cell.trim())
      .filter(cell => cell !== undefined);
  }

  /**
   * Extract HTML tables from markdown content
   */
  private static extractHtmlTables(content: string): MarkdownTable[] {
    const tables: MarkdownTable[] = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    let match;

    while ((match = tableRegex.exec(content)) !== null) {
      const tableHtml = match[1];
      const table = this.parseHtmlTable(tableHtml);
      if (table) {
        tables.push(table);
      }
    }

    return tables;
  }

  /**
   * Parse an HTML table
   */
  private static parseHtmlTable(html: string): MarkdownTable | null {
    const headers: string[] = [];
    const rows: string[][] = [];

    // Extract headers from <th> or first <tr> in <thead>
    const theadRegex = /<thead[^>]*>([\s\S]*?)<\/thead>/i;
    const theadMatch = html.match(theadRegex);
    
    if (theadMatch) {
      const thRegex = /<th[^>]*>([\s\S]*?)<\/th>/gi;
      let thMatch;
      while ((thMatch = thRegex.exec(theadMatch[1])) !== null) {
        headers.push(this.stripHtml(thMatch[1]));
      }
    }

    // If no headers found in thead, try first row
    if (headers.length === 0) {
      const firstRowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/i;
      const firstRowMatch = html.match(firstRowRegex);
      if (firstRowMatch) {
        const cellRegex = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi;
        let cellMatch;
        while ((cellMatch = cellRegex.exec(firstRowMatch[1])) !== null) {
          headers.push(this.stripHtml(cellMatch[1]));
        }
      }
    }

    // Extract data rows from <tbody> or all <tr> elements
    const tbodyRegex = /<tbody[^>]*>([\s\S]*?)<\/tbody>/i;
    const tbodyMatch = html.match(tbodyRegex);
    const rowSource = tbodyMatch ? tbodyMatch[1] : html;
    
    const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let trMatch;
    let isFirstRow = !tbodyMatch; // Skip first row if no tbody and we used it for headers
    
    while ((trMatch = trRegex.exec(rowSource)) !== null) {
      if (isFirstRow && headers.length > 0) {
        isFirstRow = false;
        continue;
      }
      
      const row: string[] = [];
      const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      let tdMatch;
      
      while ((tdMatch = tdRegex.exec(trMatch[1])) !== null) {
        row.push(this.stripHtml(tdMatch[1]));
      }
      
      if (row.length > 0) {
        rows.push(row);
      }
    }

    if (headers.length === 0 && rows.length === 0) {
      return null;
    }

    // If no headers were found, use first row as headers
    if (headers.length === 0 && rows.length > 0) {
      headers.push(...rows.shift()!);
    }

    return { headers, rows };
  }

  /**
   * Strip HTML tags from content
   */
  private static stripHtml(html: string): string {
    return html
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .trim();
  }

  /**
   * Convert a markdown table to an Excel sheet
   */
  private static tableToSheet(table: MarkdownTable, index: number): ExcelSheet {
    const name = table.title || `Table ${index + 1}`;
    
    return {
      name: this.sanitizeSheetName(name),
      headers: table.headers,
      rows: table.rows,
      config: {
        name: this.sanitizeSheetName(name),
        freezeRows: 1,
        autoFilter: true,
        alternateRows: true
      }
    };
  }

  /**
   * Sanitize sheet name to be Excel-compatible
   */
  private static sanitizeSheetName(name: string): string {
    // Excel sheet name restrictions:
    // - Max 31 characters
    // - Cannot contain: / \ ? * [ ]
    // - Cannot be empty
    
    let sanitized = name
      .replace(/[\/\\\?\*\[\]]/g, '_')
      .trim();
    
    if (sanitized.length > 31) {
      sanitized = sanitized.substring(0, 31);
    }
    
    if (sanitized.length === 0) {
      sanitized = 'Sheet';
    }
    
    return sanitized;
  }

  /**
   * Parse markdown file
   */
  static parseFile(filePath: string): ExcelSheet[] {
    const fs = require('fs');
    try {
      const content = fs.readFileSync(filePath, 'utf-8');
      return this.parse(content);
    } catch (error) {
      throw new Error(`Failed to read markdown file: ${(error as Error).message}`);
    }
  }
}