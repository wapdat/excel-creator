# 📊 Excel Creator

[![npm version](https://img.shields.io/npm/v/@knowcode/excel-creator.svg)](https://www.npmjs.com/package/@knowcode/excel-creator)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-blue.svg)](https://www.typescriptlang.org/)
[![Node.js Version](https://img.shields.io/node/v/excel-creator.svg)](https://nodejs.org)

> Create beautiful Excel spreadsheets from JSON or Markdown with rich styling, formulas, and multi-sheet support

<p align="center">
  <img src="https://via.placeholder.com/600x300/4472C4/FFFFFF?text=Excel+Creator" alt="Excel Creator Banner" />
</p>

## ✨ Features

- 📝 **Multiple Input Formats** - JSON and Markdown table support
- 📑 **Multi-sheet Workbooks** - Create complex spreadsheets with multiple tabs
- 🎨 **Rich Styling** - Professional templates with customizable colors, fonts, and borders
- 📊 **Automatic Formulas** - SUM, AVERAGE, COUNT, and more for numeric columns
- 🔧 **Highly Configurable** - Fine-tune every aspect of your spreadsheet
- 💻 **CLI & API** - Use from command line or integrate into your Node.js apps
- 🚀 **Fast & Efficient** - Built on the powerful ExcelJS library
- 📦 **TypeScript Support** - Full type definitions included

## 🚀 Quick Start

### Installation

```bash
# Global CLI installation
npm install -g @knowcode/excel-creator

# Local project installation
npm install @knowcode/excel-creator
```

### Basic CLI Usage

```bash
# Convert JSON to Excel
excel-creator --input data.json --output report.xlsx

# Parse Markdown tables
excel-creator --input README.md --format markdown --output docs.xlsx

# Use from stdin
cat data.json | excel-creator --output output.xlsx
```

### Basic API Usage

```javascript
const { createExcelFromJson, quickExcel } = require('@knowcode/excel-creator');

// Quick creation from array of objects
const data = [
  { name: 'John', age: 30, salary: 75000 },
  { name: 'Jane', age: 28, salary: 65000 },
  { name: 'Bob', age: 35, salary: 80000 }
];

await quickExcel(data, 'employees.xlsx');

// Advanced usage with configuration
await createExcelFromJson({
  sheets: [{
    name: 'Sales Report',
    headers: ['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'],
    rows: [
      ['Widget A', 100, 150, 200, 180, '=SUM(B2:E2)'],
      ['Widget B', 80, 120, 160, 200, '=SUM(B3:E3)'],
      ['Widget C', 120, 110, 140, 190, '=SUM(B4:E4)']
    ],
    config: {
      freezeRows: 1,
      autoFilter: true,
      alternateRows: true
    }
  }]
}, 'sales.xlsx');
```

## 📖 Documentation

### CLI Options

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input <path>` | Input file (JSON/Markdown) | stdin |
| `-o, --output <path>` | Output Excel file | output.xlsx |
| `-f, --format <type>` | Input format: json, markdown, auto | auto |
| `-c, --config <path>` | Configuration file for styling | - |
| `--sheet <name>` | Sheet name for single-sheet files | Sheet1 |
| `--no-formulas` | Disable automatic SUM formulas | false |
| `--no-styling` | Disable default styling | false |
| `--freeze-rows <n>` | Number of rows to freeze | 1 |
| `--alternate-rows` | Enable alternating row colors | false |
| `--currency <columns>` | Comma-separated column names for currency | - |
| `--preset <style>` | Style preset: minimal, professional, colorful, dark | professional |
| `--examples` | Show usage examples | - |
| `--debug` | Enable debug output | false |

### JSON Input Format

#### Simple Format
```json
{
  "headers": ["Name", "Age", "Department", "Salary"],
  "data": [
    ["John Doe", 30, "Engineering", 75000],
    ["Jane Smith", 28, "Marketing", 65000]
  ]
}
```

#### Multi-sheet Format
```json
{
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
}
```

### Markdown Input Format

```markdown
# Sales Report

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
```

### API Reference

#### Main Functions

##### `createExcelFromJson(input, outputPath, options)`
Create an Excel file from JSON input.

```typescript
await createExcelFromJson(
  jsonData,           // JSON string or object
  'output.xlsx',      // Output file path
  {
    addFormulas: true,     // Add automatic formulas
    applyStyling: true,    // Apply default styling
    preset: 'professional' // Style preset
  }
);
```

##### `createExcelFromMarkdown(markdown, outputPath, options)`
Create an Excel file from Markdown content.

```typescript
await createExcelFromMarkdown(
  markdownContent,    // Markdown string
  'output.xlsx',      // Output file path
  {
    addFormulas: true,
    applyStyling: true,
    preset: 'colorful'
  }
);
```

##### `quickExcel(data, outputPath, sheetName)`
Quick helper for simple conversions.

```typescript
const data = [
  { name: 'Alice', score: 95 },
  { name: 'Bob', score: 87 }
];

await quickExcel(data, 'scores.xlsx', 'Test Scores');
```

##### `createExcelBuffer(input, options)`
Create an Excel buffer without saving to file.

```typescript
const buffer = await createExcelBuffer(jsonData, {
  format: 'json',
  preset: 'dark'
});

// Use buffer for streaming, uploading, etc.
response.send(buffer);
```

#### Advanced Usage

##### Custom Styling

```javascript
import { ExcelBuilder, StyleUtils } from '@knowcode/excel-creator';

const builder = new ExcelBuilder({
  applyStyling: true,
  addFormulas: true
});

const customHeaderStyle = StyleUtils.getBrandedHeaderStyle(
  'FF0066CC',  // Primary color
  'FFFFFFFF'   // Text color
);

builder.addSheet({
  name: 'Custom Styled',
  headers: ['Column 1', 'Column 2'],
  rows: [[1, 2], [3, 4]],
  config: {
    headerStyle: customHeaderStyle,
    alternateRows: true
  }
});

await builder.save('custom.xlsx');
```

##### Formula Generation

```javascript
import { FormulaUtils } from '@knowcode/excel-creator';

// Generate various formulas
const sumFormula = FormulaUtils.sum('A2', 'A10');
const avgFormula = FormulaUtils.average('B2', 'B10');
const vlookup = FormulaUtils.vlookup('A2', 'Sheet2!A:B', 2, false);

// Use in your data
const rows = [
  [100, 200, `=${sumFormula}`],
  [150, 250, `=${avgFormula}`]
];
```

### Configuration File Format

Create a `config.json` file for reusable styling:

```json
{
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
}
```

## 🎨 Style Presets

### Professional (Default)
Clean and business-ready with gray headers and subtle borders.

### Minimal
Simple and clean without heavy formatting.

### Colorful
Vibrant blue headers with light blue alternating rows.

### Dark
Dark theme perfect for presentations and modern UIs.

## 📊 Examples

### Financial Report

```javascript
await createExcel({
  sheets: [{
    name: 'Financial Summary',
    headers: ['Category', 'Budget', 'Actual', 'Variance', '% Variance'],
    rows: [
      ['Revenue', 100000, 95000, -5000, -0.05],
      ['Expenses', 60000, 58000, 2000, 0.033],
      ['Profit', 40000, 37000, -3000, -0.075]
    ],
    config: {
      columns: [
        { key: 'Category', width: 20 },
        { key: 'Budget', format: 'currency', formula: 'sum' },
        { key: 'Actual', format: 'currency', formula: 'sum' },
        { key: 'Variance', format: 'currency', formula: 'sum' },
        { key: '% Variance', format: 'percentage' }
      ],
      freezeRows: 1,
      alternateRows: true
    }
  }]
}, 'financial-report.xlsx');
```

### Survey Results

```javascript
const surveyData = {
  sheets: [
    {
      name: 'Responses',
      headers: ['Respondent', 'Age', 'Rating', 'Feedback'],
      rows: [
        ['User001', 25, 5, 'Excellent service!'],
        ['User002', 34, 4, 'Very satisfied'],
        ['User003', 28, 5, 'Outstanding experience']
      ]
    },
    {
      name: 'Summary',
      headers: ['Metric', 'Value'],
      rows: [
        ['Total Responses', 150],
        ['Average Rating', 4.7],
        ['Satisfaction Rate', '94%']
      ]
    }
  ]
};

await createExcelFromJson(surveyData, 'survey-results.xlsx', {
  preset: 'colorful'
});
```

## 🧪 Testing

```bash
# Run tests
npm test

# Run tests with coverage
npm run test:coverage

# Run tests in watch mode
npm run test:watch
```

## 🛠️ Development

```bash
# Clone the repository
git clone https://github.com/wapdat/excel-creator.git
cd excel-creator

# Install dependencies
npm install

# Build the project
npm run build

# Run in development mode
npm run dev -- --input examples/sample.json --output test.xlsx
```

## 📄 License

MIT © Lindsay Smith

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🐛 Bug Reports

If you find a bug, please report it by opening an issue on [GitHub](https://github.com/wapdat/excel-creator/issues).

## 💖 Support

If you find this project useful, please consider:
- ⭐ Starring the repository
- 🐦 Sharing it on social media
- ☕ [Buying me a coffee](https://buymeacoffee.com/example)

## 📮 Contact

Lindsay Smith - [@wapdat](https://github.com/wapdat)

Project Link: [https://github.com/wapdat/excel-creator](https://github.com/wapdat/excel-creator)

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/wapdat">Excel Creator</a>
</p>