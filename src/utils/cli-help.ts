import chalk from 'chalk';
import Table from 'cli-table3';
import { version } from '../../package.json';

export class CliHelp {
  /**
   * Display the main help banner
   */
  static showBanner(): void {
    console.log(chalk.cyan.bold(`
╔═══════════════════════════════════════════════════════════════════╗
║                                                                   ║
║     ███████╗██╗  ██╗ ██████╗███████╗██╗                          ║
║     ██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║                          ║
║     █████╗   ╚███╔╝ ██║     █████╗  ██║                          ║
║     ██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║                          ║
║     ███████╗██╔╝ ██╗╚██████╗███████╗███████╗                     ║
║     ╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝                     ║
║                                                                   ║
║      ██████╗██████╗ ███████╗ █████╗ ████████╗ ██████╗ ██████╗    ║
║     ██╔════╝██╔══██╗██╔════╝██╔══██╗╚══██╔══╝██╔═══██╗██╔══██╗   ║
║     ██║     ██████╔╝█████╗  ███████║   ██║   ██║   ██║██████╔╝   ║
║     ██║     ██╔══██╗██╔══╝  ██╔══██║   ██║   ██║   ██║██╔══██╗   ║
║     ╚██████╗██║  ██║███████╗██║  ██║   ██║   ╚██████╔╝██║  ██║   ║
║      ╚═════╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝   ║
║                                                                   ║
╚═══════════════════════════════════════════════════════════════════╝
`));
    console.log(chalk.gray(`                         Version ${version}\n`));
    console.log(chalk.white.bold('  Create beautiful Excel spreadsheets from JSON or Markdown\n'));
  }

  /**
   * Show usage information
   */
  static showUsage(): void {
    console.log(chalk.yellow.bold('USAGE:'));
    console.log('  $ excel-creator [OPTIONS]');
    console.log('  $ excel-creator --input data.json --output report.xlsx');
    console.log('  $ cat data.json | excel-creator --output report.xlsx\n');
  }

  /**
   * Show options table
   */
  static showOptions(): void {
    console.log(chalk.yellow.bold('OPTIONS:'));
    
    const table = new Table({
      chars: {
        'top': '─', 'top-mid': '┬', 'top-left': '┌', 'top-right': '┐',
        'bottom': '─', 'bottom-mid': '┴', 'bottom-left': '└', 'bottom-right': '┘',
        'left': '│', 'left-mid': '├', 'mid': '─', 'mid-mid': '┼',
        'right': '│', 'right-mid': '┤', 'middle': '│'
      },
      style: { 'padding-left': 1, 'padding-right': 1 },
      colWidths: [25, 45, 20]
    });

    table.push(
      [chalk.cyan('Flag'), chalk.cyan('Description'), chalk.cyan('Default')],
      ['', '', ''],
      [chalk.green('-i, --input <path>'), 'Input file (JSON/Markdown)', chalk.gray('stdin')],
      [chalk.green('-o, --output <path>'), 'Output Excel file', chalk.gray('output.xlsx')],
      [chalk.green('-f, --format <type>'), 'Input format: json|markdown|auto', chalk.gray('auto')],
      [chalk.green('-c, --config <path>'), 'Configuration file for styling', chalk.gray('none')],
      [chalk.green('--sheet <name>'), 'Sheet name for single-sheet files', chalk.gray('Sheet1')],
      [chalk.green('--no-formulas'), 'Disable automatic SUM formulas', chalk.gray('false')],
      [chalk.green('--no-styling'), 'Disable default styling', chalk.gray('false')],
      [chalk.green('--freeze-rows <n>'), 'Number of rows to freeze', chalk.gray('1')],
      [chalk.green('--alternate-rows'), 'Enable alternating row colors', chalk.gray('false')],
      [chalk.green('--currency <columns>'), 'Column names for currency format', chalk.gray('none')],
      [chalk.green('--preset <style>'), 'Style preset: minimal|professional|colorful|dark', chalk.gray('professional')],
      [chalk.green('-h, --help'), 'Show help', chalk.gray('')],
      [chalk.green('-v, --version'), 'Show version', chalk.gray('')],
      [chalk.green('--examples'), 'Show usage examples', chalk.gray('')],
      [chalk.green('--debug'), 'Enable debug output', chalk.gray('false')]
    );

    console.log(table.toString());
    console.log();
  }

  /**
   * Show examples
   */
  static showExamples(): void {
    console.log(chalk.yellow.bold('EXAMPLES:\n'));
    
    const examples = [
      {
        title: 'Convert JSON to Excel with default styling',
        command: '$ excel-creator --input sales.json --output report.xlsx',
        description: 'Creates a styled Excel file from JSON data'
      },
      {
        title: 'Parse Markdown tables',
        command: '$ excel-creator --input README.md --format markdown --output docs.xlsx',
        description: 'Extracts all tables from a Markdown file'
      },
      {
        title: 'Use custom configuration',
        command: '$ excel-creator -i data.json -c style.config.json -o styled.xlsx',
        description: 'Apply custom styling from a configuration file'
      },
      {
        title: 'Process from stdin',
        command: '$ curl api.example.com/data | excel-creator --output api-data.xlsx',
        description: 'Pipe JSON data directly to Excel Creator'
      },
      {
        title: 'Multiple sheets from JSON',
        command: '$ excel-creator --input multi-sheet.json --freeze-rows 2 --alternate-rows',
        description: 'Create multi-sheet workbook with frozen headers and alternating rows'
      },
      {
        title: 'Apply currency formatting',
        command: '$ excel-creator -i financial.json --currency "Revenue,Profit" -o report.xlsx',
        description: 'Format specific columns as currency'
      },
      {
        title: 'Use style presets',
        command: '$ excel-creator --input data.json --preset dark --output dark-theme.xlsx',
        description: 'Apply a dark theme preset to the spreadsheet'
      },
      {
        title: 'Disable formulas and styling',
        command: '$ excel-creator -i raw.json --no-formulas --no-styling -o plain.xlsx',
        description: 'Create a plain Excel file without any formatting'
      }
    ];

    examples.forEach((example, index) => {
      console.log(chalk.white.bold(`${index + 1}. ${example.title}`));
      console.log(chalk.green(`   ${example.command}`));
      console.log(chalk.gray(`   ${example.description}\n`));
    });
  }

  /**
   * Show JSON input format example
   */
  static showJsonFormat(): void {
    console.log(chalk.yellow.bold('JSON INPUT FORMAT:\n'));
    console.log(chalk.gray('Simple format:'));
    console.log(chalk.green(`{
  "headers": ["Name", "Age", "Department", "Salary"],
  "data": [
    ["John Doe", 30, "Engineering", 75000],
    ["Jane Smith", 28, "Marketing", 65000],
    ["Bob Johnson", 35, "Sales", 80000]
  ]
}`));

    console.log(chalk.gray('\nMulti-sheet format:'));
    console.log(chalk.green(`{
  "sheets": [
    {
      "name": "Employees",
      "headers": ["Name", "Department", "Salary"],
      "rows": [
        ["John Doe", "Engineering", 75000],
        ["Jane Smith", "Marketing", 65000]
      ],
      "config": {
        "freezeRows": 1,
        "autoFilter": true,
        "alternateRows": true
      }
    },
    {
      "name": "Departments",
      "headers": ["Department", "Budget", "Head Count"],
      "rows": [
        ["Engineering", 500000, 25],
        ["Marketing", 300000, 15]
      ]
    }
  ]
}`));
  }

  /**
   * Show markdown format example
   */
  static showMarkdownFormat(): void {
    console.log(chalk.yellow.bold('MARKDOWN INPUT FORMAT:\n'));
    console.log(chalk.gray('Tables in your markdown will be automatically detected:'));
    console.log(chalk.green(`# Sales Report

## Q1 Results

| Product | Units Sold | Revenue |
|---------|------------|---------|
| Widget A | 150 | $7,500 |
| Widget B | 200 | $12,000 |
| Widget C | 100 | $8,500 |

## Q2 Results

| Product | Units Sold | Revenue |
|---------|------------|---------|
| Widget A | 180 | $9,000 |
| Widget B | 220 | $13,200 |
| Widget C | 130 | $11,050 |
`));
  }

  /**
   * Show configuration format
   */
  static showConfigFormat(): void {
    console.log(chalk.yellow.bold('CONFIGURATION FILE FORMAT:\n'));
    console.log(chalk.green(`{
  "freezeRows": 2,
  "autoFilter": true,
  "alternateRows": true,
  "headerStyle": {
    "font": {
      "bold": true,
      "size": 14,
      "color": { "argb": "FFFFFFFF" }
    },
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": { "argb": "FF4472C4" }
    }
  },
  "columns": [
    {
      "key": "Revenue",
      "format": "currency",
      "width": 15,
      "formula": "sum"
    },
    {
      "key": "Growth",
      "format": "percentage",
      "width": 12
    }
  ]
}`));
  }

  /**
   * Show error message
   */
  static showError(message: string, suggestion?: string): void {
    console.log(chalk.red.bold('\n✖ Error:'), chalk.red(message));
    
    if (suggestion) {
      console.log(chalk.yellow('\n💡 Suggestion:'), chalk.white(suggestion));
    }
    
    console.log(chalk.gray('\nFor more help, run: excel-creator --help\n'));
  }

  /**
   * Show success message
   */
  static showSuccess(outputPath: string, stats?: { sheets: number; rows: number }): void {
    console.log(chalk.green.bold('\n✓ Success!'), chalk.white(`Excel file created: ${outputPath}`));
    
    if (stats) {
      console.log(chalk.gray(`  → ${stats.sheets} sheet(s), ${stats.rows} row(s) processed\n`));
    }
  }

  /**
   * Show progress spinner text
   */
  static getSpinnerText(action: string): string {
    return chalk.cyan(`${action}...`);
  }

  /**
   * Show debug information
   */
  static showDebug(label: string, data: any): void {
    console.log(chalk.magenta.bold(`[DEBUG] ${label}:`));
    console.log(chalk.gray(JSON.stringify(data, null, 2)));
  }

  /**
   * Show footer with links
   */
  static showFooter(): void {
    console.log(chalk.gray('\n─────────────────────────────────────────────────\n'));
    console.log(chalk.white.bold('MORE INFO:'));
    console.log(chalk.blue('  Documentation: https://github.com/wapdat/excel-creator'));
    console.log(chalk.blue('  Issues: https://github.com/wapdat/excel-creator/issues'));
    console.log(chalk.gray('\n  Made with ❤️  by Excel Creator\n'));
  }
}