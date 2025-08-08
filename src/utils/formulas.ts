export class FormulaUtils {
  /**
   * Generate a SUM formula for a column range
   */
  static sum(startCell: string, endCell: string): string {
    return `SUM(${startCell}:${endCell})`;
  }

  /**
   * Generate an AVERAGE formula for a column range
   */
  static average(startCell: string, endCell: string): string {
    return `AVERAGE(${startCell}:${endCell})`;
  }

  /**
   * Generate a COUNT formula for a column range
   */
  static count(startCell: string, endCell: string): string {
    return `COUNT(${startCell}:${endCell})`;
  }

  /**
   * Generate a COUNTA formula (counts non-empty cells)
   */
  static countA(startCell: string, endCell: string): string {
    return `COUNTA(${startCell}:${endCell})`;
  }

  /**
   * Generate a MIN formula for a column range
   */
  static min(startCell: string, endCell: string): string {
    return `MIN(${startCell}:${endCell})`;
  }

  /**
   * Generate a MAX formula for a column range
   */
  static max(startCell: string, endCell: string): string {
    return `MAX(${startCell}:${endCell})`;
  }

  /**
   * Generate a MEDIAN formula for a column range
   */
  static median(startCell: string, endCell: string): string {
    return `MEDIAN(${startCell}:${endCell})`;
  }

  /**
   * Generate a STDEV formula for a column range
   */
  static stdev(startCell: string, endCell: string): string {
    return `STDEV(${startCell}:${endCell})`;
  }

  /**
   * Generate a VAR formula for a column range
   */
  static variance(startCell: string, endCell: string): string {
    return `VAR(${startCell}:${endCell})`;
  }

  /**
   * Generate a CONCATENATE formula
   */
  static concatenate(...cells: string[]): string {
    return `CONCATENATE(${cells.join(',')})`;
  }

  /**
   * Generate an IF formula
   */
  static if(condition: string, trueValue: string, falseValue: string): string {
    return `IF(${condition},${trueValue},${falseValue})`;
  }

  /**
   * Generate a VLOOKUP formula
   */
  static vlookup(lookupValue: string, tableArray: string, colIndex: number, rangeLookup: boolean = false): string {
    return `VLOOKUP(${lookupValue},${tableArray},${colIndex},${rangeLookup ? 'TRUE' : 'FALSE'})`;
  }

  /**
   * Generate a HLOOKUP formula
   */
  static hlookup(lookupValue: string, tableArray: string, rowIndex: number, rangeLookup: boolean = false): string {
    return `HLOOKUP(${lookupValue},${tableArray},${rowIndex},${rangeLookup ? 'TRUE' : 'FALSE'})`;
  }

  /**
   * Generate a SUMIF formula
   */
  static sumIf(range: string, criteria: string, sumRange?: string): string {
    if (sumRange) {
      return `SUMIF(${range},"${criteria}",${sumRange})`;
    }
    return `SUMIF(${range},"${criteria}")`;
  }

  /**
   * Generate a COUNTIF formula
   */
  static countIf(range: string, criteria: string): string {
    return `COUNTIF(${range},"${criteria}")`;
  }

  /**
   * Generate a SUMPRODUCT formula
   */
  static sumProduct(...ranges: string[]): string {
    return `SUMPRODUCT(${ranges.join(',')})`;
  }

  /**
   * Generate a percentage formula
   */
  static percentage(numerator: string, denominator: string): string {
    return `${numerator}/${denominator}`;
  }

  /**
   * Generate a growth rate formula
   */
  static growthRate(newValue: string, oldValue: string): string {
    return `(${newValue}-${oldValue})/${oldValue}`;
  }

  /**
   * Convert column letter to number (A -> 1, B -> 2, etc.)
   */
  static columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * Convert number to column letter (1 -> A, 2 -> B, etc.)
   */
  static numberToColumn(num: number): string {
    let column = '';
    let n = num;
    
    while (n > 0) {
      n--;
      column = String.fromCharCode(65 + (n % 26)) + column;
      n = Math.floor(n / 26);
    }
    
    return column;
  }

  /**
   * Get cell reference (e.g., A1, B2)
   */
  static getCellReference(column: number | string, row: number): string {
    const col = typeof column === 'number' ? this.numberToColumn(column) : column;
    return `${col}${row}`;
  }

  /**
   * Get range reference (e.g., A1:B10)
   */
  static getRangeReference(
    startColumn: number | string,
    startRow: number,
    endColumn: number | string,
    endRow: number
  ): string {
    const startCell = this.getCellReference(startColumn, startRow);
    const endCell = this.getCellReference(endColumn, endRow);
    return `${startCell}:${endCell}`;
  }

  /**
   * Generate formula based on type
   */
  static generateFormula(
    type: 'sum' | 'average' | 'count' | 'min' | 'max' | 'median' | 'stdev',
    startCell: string,
    endCell: string
  ): string {
    switch (type) {
      case 'sum':
        return this.sum(startCell, endCell);
      case 'average':
        return this.average(startCell, endCell);
      case 'count':
        return this.count(startCell, endCell);
      case 'min':
        return this.min(startCell, endCell);
      case 'max':
        return this.max(startCell, endCell);
      case 'median':
        return this.median(startCell, endCell);
      case 'stdev':
        return this.stdev(startCell, endCell);
      default:
        return this.sum(startCell, endCell);
    }
  }

  /**
   * Check if a value should have a formula applied
   */
  static isNumericColumn(values: any[]): boolean {
    if (values.length === 0) return false;
    
    const nonEmptyValues = values.filter(v => v != null && v !== '');
    if (nonEmptyValues.length === 0) return false;
    
    return nonEmptyValues.every(value => {
      if (typeof value === 'number') return true;
      if (typeof value === 'string') {
        // Remove currency symbols and check if it's a number
        const cleaned = value.replace(/[$£€¥,]/g, '').trim();
        return !isNaN(Number(cleaned)) && cleaned !== '';
      }
      return false;
    });
  }

  /**
   * Generate subtotal formula for filtered data
   */
  static subtotal(functionNum: number, range: string): string {
    // Function numbers:
    // 9 = SUM, 1 = AVERAGE, 2 = COUNT, 3 = COUNTA,
    // 4 = MAX, 5 = MIN, 6 = PRODUCT, 7 = STDEV, 8 = STDEVP,
    // 9 = SUM, 10 = VAR, 11 = VARP
    return `SUBTOTAL(${functionNum},${range})`;
  }

  /**
   * Generate a formula for running total
   */
  static runningTotal(startCell: string, currentCell: string): string {
    return `SUM($${startCell}:${currentCell})`;
  }

  /**
   * Generate a formula for percentage of total
   */
  static percentageOfTotal(cellValue: string, totalCell: string): string {
    return `${cellValue}/$${totalCell}`;
  }

  /**
   * Generate a formula for year-over-year growth
   */
  static yearOverYearGrowth(currentYear: string, previousYear: string): string {
    return `(${currentYear}-${previousYear})/${previousYear}`;
  }

  /**
   * Generate a formula for compound annual growth rate (CAGR)
   */
  static cagr(endValue: string, startValue: string, periods: number): string {
    return `POWER(${endValue}/${startValue},1/${periods})-1`;
  }
}