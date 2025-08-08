const ExcelJS = require('exceljs');
const fs = require('fs');

async function createStreamingExcel() {
  console.log('Testing streaming Excel creation for better compatibility...\n');

  // Test 1: Using streaming workbook writer
  console.log('1. Creating file with streaming API...');
  const options = {
    filename: 'stream-test-1.xlsx',
    useStyles: false,  // Disable styles for compatibility
    useSharedStrings: false  // Disable shared strings for compatibility
  };

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('Sheet1');

  worksheet.addRow(['Name', 'Value']);
  worksheet.addRow(['Item 1', 100]);
  worksheet.addRow(['Item 2', 200]);
  
  await worksheet.commit();
  await workbook.commit();
  console.log('   Created: stream-test-1.xlsx');

  // Test 2: Write buffer then save
  console.log('\n2. Creating file via buffer...');
  const wb2 = new ExcelJS.Workbook();
  const ws2 = wb2.addWorksheet('Data');
  
  ws2.addRow(['Product', 'Price']);
  ws2.addRow(['Apple', 1.50]);
  ws2.addRow(['Banana', 0.75]);
  
  const buffer = await wb2.xlsx.writeBuffer();
  fs.writeFileSync('buffer-test-2.xlsx', buffer);
  console.log('   Created: buffer-test-2.xlsx');

  // Test 3: Set workbook properties that might affect Mac
  console.log('\n3. Creating file with explicit properties...');
  const wb3 = new ExcelJS.Workbook();
  
  // Set properties that might help with Mac compatibility
  wb3.creator = 'Excel Creator';
  wb3.lastModifiedBy = 'Excel Creator';
  wb3.created = new Date();
  wb3.modified = new Date();
  wb3.properties = {
    title: 'Test Workbook',
    author: 'Excel Creator'
  };
  
  // Explicitly set calculation properties
  wb3.calcProperties = {
    fullCalcOnLoad: false
  };
  
  const ws3 = wb3.addWorksheet('Test', {
    properties: {
      defaultRowHeight: 15,
      defaultColWidth: 10
    }
  });
  
  ws3.addRow(['A', 'B']);
  ws3.addRow([1, 2]);
  ws3.addRow([3, 4]);
  
  await wb3.xlsx.writeFile('properties-test-3.xlsx');
  console.log('   Created: properties-test-3.xlsx');

  console.log('\n✅ Testing complete. Try these files in Mac Excel.');
}

createStreamingExcel().catch(console.error);