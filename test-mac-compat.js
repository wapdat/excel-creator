const ExcelJS = require('exceljs');

async function createMacCompatibleExcel() {
  console.log('Creating Mac-compatible Excel files...\n');

  // Test 1: Absolute minimal Excel file
  console.log('1. Creating absolute minimal file...');
  const wb1 = new ExcelJS.Workbook();
  const ws1 = wb1.addWorksheet('Sheet1');
  
  ws1.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Value', key: 'value', width: 15 }
  ];
  
  ws1.addRow({ name: 'Item 1', value: 100 });
  ws1.addRow({ name: 'Item 2', value: 200 });
  
  await wb1.xlsx.writeFile('mac-test-1-minimal.xlsx');
  console.log('   Created: mac-test-1-minimal.xlsx');

  // Test 2: Without column definitions
  console.log('\n2. Creating file without column definitions...');
  const wb2 = new ExcelJS.Workbook();
  const ws2 = wb2.addWorksheet('Data');
  
  // Add rows directly without column setup
  ws2.addRow(['Product', 'Price']);
  ws2.addRow(['Apple', 1.50]);
  ws2.addRow(['Banana', 0.75]);
  
  await wb2.xlsx.writeFile('mac-test-2-simple.xlsx');
  console.log('   Created: mac-test-2-simple.xlsx');

  // Test 3: With basic styling only
  console.log('\n3. Creating file with minimal styling...');
  const wb3 = new ExcelJS.Workbook();
  const ws3 = wb3.addWorksheet('Styled');
  
  const headerRow = ws3.addRow(['Name', 'Amount']);
  headerRow.font = { bold: true };
  
  ws3.addRow(['Sales', 5000]);
  ws3.addRow(['Costs', 3000]);
  
  await wb3.xlsx.writeFile('mac-test-3-styled.xlsx');
  console.log('   Created: mac-test-3-styled.xlsx');

  // Test 4: With numbers only (no formulas)
  console.log('\n4. Creating file with numbers only...');
  const wb4 = new ExcelJS.Workbook();
  const ws4 = wb4.addWorksheet('Numbers');
  
  ws4.addRow(['Month', 'Revenue', 'Expenses']);
  ws4.addRow(['Jan', 10000, 7000]);
  ws4.addRow(['Feb', 12000, 8000]);
  ws4.addRow(['Mar', 15000, 9000]);
  
  await wb4.xlsx.writeFile('mac-test-4-numbers.xlsx');
  console.log('   Created: mac-test-4-numbers.xlsx');

  // Test 5: Check our package's minimal output
  console.log('\n5. Testing our package with NO features...');
  const { ExcelBuilder } = require('./dist/index');
  
  const builder = new ExcelBuilder({
    addFormulas: false,
    applyStyling: false
  });
  
  builder.addSheet({
    name: 'Test',
    headers: ['A', 'B', 'C'],
    rows: [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9]
    ]
  });
  
  await builder.save('mac-test-5-package.xlsx');
  console.log('   Created: mac-test-5-package.xlsx');

  console.log('\n✅ Done! Please try opening these files in Excel on Mac.');
  console.log('Start with mac-test-1-minimal.xlsx - if that doesn\'t work, the issue is with ExcelJS itself.');
}

createMacCompatibleExcel().catch(console.error);