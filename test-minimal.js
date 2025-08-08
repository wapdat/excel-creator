const ExcelJS = require('exceljs');

async function createMinimalExcel() {
  console.log('Creating minimal Excel file for testing...');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Test');
  
  // Add simple data without any formulas or complex formatting
  worksheet.addRow(['Name', 'Value']);
  worksheet.addRow(['Item 1', 100]);
  worksheet.addRow(['Item 2', 200]);
  
  await workbook.xlsx.writeFile('minimal-test.xlsx');
  console.log('Created minimal-test.xlsx');
  
  // Now test with our package
  const { ExcelBuilder } = require('./dist/index');
  
  const builder = new ExcelBuilder({
    addFormulas: false,  // Disable formulas
    applyStyling: false  // Disable styling
  });
  
  builder.addSheet({
    name: 'Simple Test',
    headers: ['Product', 'Price'],
    rows: [
      ['Apple', 1.50],
      ['Banana', 0.75],
      ['Orange', 2.00]
    ]
  });
  
  await builder.save('simple-test.xlsx');
  console.log('Created simple-test.xlsx with our package');
}

createMinimalExcel().catch(console.error);