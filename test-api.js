const { quickExcel, createExcelFromJson } = require('./dist/index');

async function testAPI() {
  console.log('Testing Excel Creator API...\n');

  // Test 1: Quick Excel
  console.log('1. Testing quickExcel...');
  const data = [
    { name: 'Alice', score: 95, grade: 'A' },
    { name: 'Bob', score: 87, grade: 'B' },
    { name: 'Charlie', score: 92, grade: 'A' }
  ];
  
  await quickExcel(data, 'test-quick.xlsx', 'Grades');
  console.log('   ✓ Created test-quick.xlsx\n');

  // Test 2: Create from JSON with configuration
  console.log('2. Testing createExcelFromJson...');
  const jsonData = {
    sheets: [{
      name: 'Sales',
      headers: ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
      rows: [
        ['Widget A', 1000, 1200, 1400, 1600],
        ['Widget B', 800, 900, 1000, 1100]
      ],
      config: {
        freezeRows: 1,
        autoFilter: true,
        columns: [
          { key: 'Product', width: 20 },
          { key: 'Q1', format: 'currency', formula: 'sum' },
          { key: 'Q2', format: 'currency', formula: 'sum' },
          { key: 'Q3', format: 'currency', formula: 'sum' },
          { key: 'Q4', format: 'currency', formula: 'sum' }
        ]
      }
    }]
  };
  
  await createExcelFromJson(jsonData, 'test-configured.xlsx', { preset: 'colorful' });
  console.log('   ✓ Created test-configured.xlsx\n');

  console.log('All tests passed! 🎉');
}

testAPI().catch(console.error);