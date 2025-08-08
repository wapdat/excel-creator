const { createExcelFromJson, ExcelBuilder } = require('./dist/index');

async function testComplete() {
  console.log('Testing complete Excel generation with formulas...\n');

  // Test 1: Simple sheet with formulas
  console.log('1. Creating sheet with working formulas...');
  const builder1 = new ExcelBuilder({
    addFormulas: true,
    applyStyling: true
  });

  builder1.addSheet({
    name: 'Sales Report',
    headers: ['Month', 'Revenue', 'Expenses', 'Profit'],
    rows: [
      ['January', 10000, 7000, 3000],
      ['February', 12000, 8000, 4000],
      ['March', 15000, 9000, 6000],
      ['April', 11000, 7500, 3500],
      ['May', 13000, 8500, 4500],
      ['June', 14000, 9000, 5000]
    ],
    config: {
      freezeRows: 1,
      autoFilter: true,
      alternateRows: true,
      columns: [
        { key: 'Month', width: 15 },
        { key: 'Revenue', format: 'currency', formula: 'sum', width: 15 },
        { key: 'Expenses', format: 'currency', formula: 'sum', width: 15 },
        { key: 'Profit', format: 'currency', formula: 'sum', width: 15 }
      ]
    }
  });

  await builder1.save('test-complete-formulas.xlsx');
  console.log('   ✓ Created test-complete-formulas.xlsx\n');

  // Test 2: Multiple sheets
  console.log('2. Creating multi-sheet workbook...');
  const multiSheetData = {
    sheets: [
      {
        name: 'Q1 Sales',
        headers: ['Product', 'Units', 'Price', 'Total'],
        rows: [
          ['Laptop', 50, 999.99, 49999.50],
          ['Mouse', 200, 29.99, 5998.00],
          ['Keyboard', 150, 79.99, 11998.50],
          ['Monitor', 75, 299.99, 22499.25]
        ],
        config: {
          freezeRows: 1,
          columns: [
            { key: 'Product', width: 20 },
            { key: 'Units', format: 'number', formula: 'sum', width: 12 },
            { key: 'Price', format: 'currency', width: 12 },
            { key: 'Total', format: 'currency', formula: 'sum', width: 15 }
          ]
        }
      },
      {
        name: 'Q2 Sales',
        headers: ['Product', 'Units', 'Price', 'Total'],
        rows: [
          ['Laptop', 65, 999.99, 64999.35],
          ['Mouse', 250, 29.99, 7497.50],
          ['Keyboard', 180, 79.99, 14398.20],
          ['Monitor', 90, 299.99, 26999.10]
        ],
        config: {
          freezeRows: 1,
          columns: [
            { key: 'Product', width: 20 },
            { key: 'Units', format: 'number', formula: 'sum', width: 12 },
            { key: 'Price', format: 'currency', width: 12 },
            { key: 'Total', format: 'currency', formula: 'sum', width: 15 }
          ]
        }
      }
    ]
  };

  await createExcelFromJson(multiSheetData, 'test-complete-multisheet.xlsx', {
    preset: 'professional'
  });
  console.log('   ✓ Created test-complete-multisheet.xlsx\n');

  // Test 3: Different style presets
  console.log('3. Testing style presets...');
  const presets = ['minimal', 'professional', 'colorful', 'dark'];
  
  for (const preset of presets) {
    const builder = new ExcelBuilder({
      addFormulas: true,
      applyStyling: true
    });

    const presetData = {
      name: `${preset} Style`,
      headers: ['Item', 'Quantity', 'Price'],
      rows: [
        ['Item A', 10, 25.50],
        ['Item B', 20, 15.75],
        ['Item C', 15, 30.00]
      ],
      config: {
        freezeRows: 1,
        columns: [
          { key: 'Item', width: 20 },
          { key: 'Quantity', format: 'number', formula: 'sum', width: 12 },
          { key: 'Price', format: 'currency', formula: 'average', width: 12 }
        ]
      }
    };

    await createExcelFromJson(
      { sheets: [presetData] },
      `test-complete-${preset}.xlsx`,
      { preset }
    );
    console.log(`   ✓ Created test-complete-${preset}.xlsx`);
  }

  console.log('\n✅ All tests completed successfully!');
  console.log('Please open the generated Excel files to verify they work correctly.');
}

testComplete().catch(console.error);