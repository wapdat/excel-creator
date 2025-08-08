const ExcelJS = require('exceljs');

async function testFormulas() {
  console.log('Testing formula syntax...');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Formulas');
  
  // Add headers
  worksheet.addRow(['Item', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']);
  
  // Add data
  worksheet.addRow(['Product A', 100, 200, 300, 400, null]);
  worksheet.addRow(['Product B', 150, 250, 350, 450, null]);
  
  // Test different formula syntaxes
  console.log('Testing formula assignment methods...');
  
  // Method 1: Direct formula string (CORRECT)
  worksheet.getCell('F2').value = { formula: '=SUM(B2:E2)' };
  worksheet.getCell('F3').value = { formula: '=SUM(B3:E3)' };
  
  // Add total row
  const totalRow = worksheet.addRow(['Total', null, null, null, null, null]);
  worksheet.getCell(`B${totalRow.number}`).value = { formula: '=SUM(B2:B3)' };
  worksheet.getCell(`C${totalRow.number}`).value = { formula: '=SUM(C2:C3)' };
  worksheet.getCell(`D${totalRow.number}`).value = { formula: '=SUM(D2:D3)' };
  worksheet.getCell(`E${totalRow.number}`).value = { formula: '=SUM(E2:E3)' };
  worksheet.getCell(`F${totalRow.number}`).value = { formula: '=SUM(F2:F3)' };
  
  await workbook.xlsx.writeFile('test-formulas.xlsx');
  console.log('Created test-formulas.xlsx with formulas');
  
  // Now let's check what our package is generating
  const { FormulaUtils } = require('./dist/index');
  
  const formula = FormulaUtils.sum('B2', 'E2');
  console.log('Our formula generator returns:', formula);
  console.log('Should be used as: { formula: "=' + formula + '" }');
}

testFormulas().catch(console.error);