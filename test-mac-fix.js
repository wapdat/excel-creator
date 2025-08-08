const { ExcelBuilder } = require('./dist/index');

async function testMacFix() {
  console.log('Testing Mac Excel compatibility fixes...\n');

  // Test 1: No styling, no formulas, no features
  console.log('1. Creating bare minimum Excel...');
  const builder1 = new ExcelBuilder({
    addFormulas: false,
    applyStyling: false
  });

  builder1.addSheet({
    name: 'Simple',
    headers: ['Name', 'Value'],
    rows: [
      ['Item A', 100],
      ['Item B', 200],
      ['Item C', 300]
    ]
  });

  await builder1.save('mac-fixed-1-bare.xlsx');
  console.log('   ✓ Created: mac-fixed-1-bare.xlsx');

  // Test 2: With formulas but no styling
  console.log('\n2. Creating Excel with formulas only...');
  const builder2 = new ExcelBuilder({
    addFormulas: true,
    applyStyling: false
  });

  builder2.addSheet({
    name: 'Formulas',
    headers: ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
    rows: [
      ['Widget A', 1000, 1200, 1400, 1600],
      ['Widget B', 800, 900, 1000, 1100],
      ['Widget C', 600, 700, 800, 900]
    ]
  });

  await builder2.save('mac-fixed-2-formulas.xlsx');
  console.log('   ✓ Created: mac-fixed-2-formulas.xlsx');

  // Test 3: With minimal styling
  console.log('\n3. Creating Excel with minimal styling...');
  const builder3 = new ExcelBuilder({
    addFormulas: false,
    applyStyling: true
  });

  builder3.addSheet({
    name: 'Styled',
    headers: ['Name', 'Amount', 'Status'],
    rows: [
      ['Sales', 5000, 'Complete'],
      ['Returns', 500, 'Pending'],
      ['Net', 4500, 'Final']
    ],
    config: {
      name: 'Styled',
      alternateRows: true
    }
  });

  await builder3.save('mac-fixed-3-styled.xlsx');
  console.log('   ✓ Created: mac-fixed-3-styled.xlsx');

  // Test 4: Without problematic features
  console.log('\n4. Creating Excel without freeze/filter...');
  const builder4 = new ExcelBuilder({
    addFormulas: true,
    applyStyling: true
  });

  builder4.addSheet({
    name: 'NoFreeze',
    headers: ['Month', 'Revenue', 'Expenses'],
    rows: [
      ['Jan', 10000, 7000],
      ['Feb', 12000, 8000],
      ['Mar', 15000, 9000]
    ],
    config: {
      name: 'NoFreeze',
      // NO freezeRows
      // NO autoFilter
      alternateRows: true,
      columns: [
        { key: 'Month', width: 15 },
        { key: 'Revenue', format: 'currency', formula: 'sum', width: 15 },
        { key: 'Expenses', format: 'currency', formula: 'sum', width: 15 }
      ]
    }
  });

  await builder4.save('mac-fixed-4-nofreeze.xlsx');
  console.log('   ✓ Created: mac-fixed-4-nofreeze.xlsx');

  console.log('\n✅ Testing complete!');
  console.log('Please try opening these files in Mac Excel.');
  console.log('Start with mac-fixed-1-bare.xlsx for the simplest test.');
}

testMacFix().catch(console.error);