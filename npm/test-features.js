// Tests for the full feature set mirrored from the Dart API.
// Each feature is written, saved to .xlsx bytes and read back.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { Excel } = require('./index.js');

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64'
);

let failed = 0;
function test(name, fn) {
  try {
    fn();
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}\n      ${(e && e.message) || e}`);
  }
}

const reload = (wb) => Excel.read(wb.encode());

function salesSheet(wb, name = 'Sales') {
  const s = wb.sheet(name);
  s.appendRow(['Region', 'Units', 'Price', 'Date']);
  s.appendRow(['North', 10, 2.5, new Date(2026, 0, 15)]);
  s.appendRow(['South', 7, 3, new Date(2026, 1, 20)]);
  s.appendRow(['East', 4, 1.5, new Date(2026, 2, 5)]);
  return s;
}

console.log('Cells');
test('setStyle merges with the existing style; resetStyle clears it', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').value = 1234.5;
  s.cell('A1').setStyle({ numberFormat: '$#,##0.00' });
  s.cell('A1').setStyle({ bold: true });
  const st = s.cell('A1').style;
  assert.strictEqual(st.bold, true);
  assert.strictEqual(st.numberFormat, '$#,##0.00');
  assert.strictEqual(s.cell('A1').displayText, '$1,234.50');
  s.cell('A1').resetStyle();
  assert.strictEqual(s.cell('A1').style.bold, false);
});

test('diagonal borders, wrap, shrink, locked and hidden survive save', () => {
  const wb = Excel.create();
  const c = wb.sheet('Sheet1').cell('B2');
  c.value = 'x';
  c.setStyle({
    diagonalBorder: { style: 'thin', color: '#FF0000' },
    diagonalUp: true,
    wrapText: true,
    locked: false,
    hidden: true,
  });
  const st = reload(wb).sheet('Sheet1').cell('B2').style;
  assert.deepStrictEqual([st.diagonalUp, st.wrapText, st.locked, st.hidden], [true, true, false, true]);
  assert.strictEqual(st.diagonalBorder.style, 'thin');
});

test('setTime, dateValue and time values', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').setTime('14:30:05');
  s.cell('A2').setTime(8, 15);
  s.cell('A3').value = new Date(2026, 9, 4);
  const back = reload(wb).sheet('Sheet1');
  assert.deepStrictEqual([back.cell('A1').type, back.cell('A1').value], ['time', '14:30:05']);
  assert.strictEqual(back.cell('A2').value, '08:15:00');
  const d = back.cell('A3').dateValue;
  assert.ok(d instanceof Date);
  assert.deepStrictEqual([d.getFullYear(), d.getMonth(), d.getDate()], [2026, 9, 4]);
});

test('cachedValue of formulas read from a file', () => {
  const file = path.join(__dirname, '..', 'test', 'test_resources', 'example.xlsx');
  if (!fs.existsSync(file)) return;
  const wb = Excel.fromFile(file);
  for (const name of wb.sheets) {
    for (const row of wb.sheet(name).rows) {
      for (const cell of row) if (cell && cell.type === 'formula') return void cell.cachedValue;
    }
  }
});

test('hyperlinks: url, email, cell, location, range and clear', () => {
  const wb = Excel.create();
  const s = wb.sheet('Links');
  wb.createSheet('Q1 Sales');
  s.cell('A1').setHyperlink({ url: 'https://pub.dev', tooltip: 'pub' });
  s.cell('A2').setHyperlink({ email: 'sales@example.com', subject: 'Q3 report' }, { text: 'Contact sales' });
  s.cell('A3').setHyperlink({ sheet: 'Q1 Sales', cell: 'B4' }, { text: 'Go to Q1' });
  s.setHyperlinkRange('A5:C5', { location: 'Links!A1' });
  const back = reload(wb).sheet('Links');
  assert.strictEqual(back.cell('A1').getHyperlink().url, 'https://pub.dev');
  assert.strictEqual(back.cell('A2').getHyperlink().url, 'mailto:sales@example.com?subject=Q3%20report');
  assert.strictEqual(back.cell('A2').value, 'Contact sales');
  assert.strictEqual(back.cell('A3').getHyperlink().location, "'Q1 Sales'!B4");
  assert.ok(Object.keys(back.hyperlinks).includes('A5:C5'));
  s.clearHyperlinks();
  assert.deepStrictEqual(Object.keys(s.hyperlinks), []);
});

console.log('Rows, columns and ranges');
test('insert / remove rows and columns', () => {
  const s = salesSheet(Excel.create());
  s.insertRow(1);
  assert.strictEqual(s.cell('A3').value, 'North');
  s.removeRow(1);
  assert.strictEqual(s.cell('A2').value, 'North');
  s.insertColumn(0);
  assert.strictEqual(s.cell('B1').value, 'Region');
  s.removeColumn(0);
  assert.strictEqual(s.cell('A1').value, 'Region');
});

test('insertRowIterables, clearRow and rangeValues', () => {
  const s = Excel.create().sheet('Sheet1');
  s.insertRowIterables(['a', 'b', 'c'], 2, { startingColumn: 1 });
  assert.deepStrictEqual(s.rangeValues('B3:D3'), [['a', 'b', 'c']]);
  assert.strictEqual(s.clearRow(2), true);
  assert.deepStrictEqual(s.rangeValues('B3:D3'), [[null, null, null]]);
});

test('merge with a custom value', () => {
  const s = Excel.create().sheet('Sheet1');
  s.merge('A1', 'C1', 'Banner');
  assert.strictEqual(s.cell('A1').value, 'Banner');
  assert.deepStrictEqual(s.spannedItems, ['A1:C1']);
});

test('findAndReplace with strings and RegExp', () => {
  const s = Excel.create().sheet('Sheet1');
  s.appendRow(['Flutter loves Excel', 'FLUTTER']);
  assert.strictEqual(s.findAndReplace('Flutter', 'Dart'), 1);
  assert.strictEqual(s.cell('A1').value, 'Dart loves Excel');
  assert.strictEqual(s.findAndReplace(/flutter/i, 'Dart'), 1);
  assert.strictEqual(s.cell('B1').value, 'Dart');
});

console.log('Export & import');
test('rowsAsMaps, rowsAsValues, toJson, toCsv', () => {
  const wb = Excel.create();
  const s = salesSheet(wb);
  const maps = s.rowsAsMaps();
  assert.strictEqual(maps.length, 3);
  assert.deepStrictEqual([maps[0].Region, maps[0].Units, maps[0].Price], ['North', 10, 2.5]);
  assert.ok(maps[0].Date instanceof Date);
  assert.strictEqual(s.rowsAsMaps({ mode: 'displayText' })[0].Units, '10');
  assert.deepStrictEqual(s.rowsAsValues()[0], ['Region', 'Units', 'Price', 'Date']);
  assert.strictEqual(JSON.parse(s.toJson())[1].Region, 'South');
  // Decimals get Dart's default '0.00' number format, so the CSV shows 2.50.
  assert.ok(s.toCsv({ separator: ';' }).startsWith('Region;Units;Price;Date\r\nNorth;10;2.50;'));
  assert.deepStrictEqual(Object.keys(wb.toMaps()), ['Sales']);
  assert.strictEqual(JSON.parse(wb.toJson()).Sales.length, 3);
});

test('appendRowsFromMaps writes the header and matches columns', () => {
  const s = Excel.create().sheet('Imported');
  s.appendRowsFromMaps([{ Name: 'Eva', Age: 40, Joined: new Date(2026, 0, 15) }]);
  s.appendRowsFromMaps([{ Age: 31, Name: 'Ana' }]);
  assert.deepStrictEqual(s.rowsAsValues()[0], ['Name', 'Age', 'Joined']);
  assert.deepStrictEqual(s.rowsAsValues()[2].slice(0, 2), ['Ana', 31]);
});

console.log('Conditional formatting');
test('cellIs, between, containsText, expression, duplicates and unique', () => {
  const wb = Excel.create();
  const s = salesSheet(wb);
  s.addConditionalFormatting('B2:B4', {
    type: 'cellIs',
    operator: 'greaterThan',
    value: 5,
    style: { backgroundColor: '#C6EFCE', fontColor: '#006100', bold: true },
  });
  s.addConditionalFormatting('C2:C4', [
    { type: 'cellIs', operator: 'between', value: 1, value2: 2, style: { fontColor: '#9C0006' } },
    { type: 'expression', formula: '=$C2>2.5', style: { italic: true }, priority: 2 },
  ]);
  s.addConditionalFormatting('A2:A4', { type: 'containsText', text: 'th', style: { underline: 'single' } });
  s.addConditionalFormatting('A2:A4', { type: 'duplicateValues', style: { backgroundColor: '#FFC7CE' } });
  s.addConditionalFormatting('A2:A4', { type: 'uniqueValues', style: { strikethrough: true } });
  const back = reload(wb).sheet('Sales').conditionalFormattings;
  assert.strictEqual(back.length, 5);
  assert.deepStrictEqual(back[0].rules[0].formulae, ['5']);
  assert.deepStrictEqual(back[1].rules.map((r) => r.type), ['cellIs', 'expression']);
  s.clearConditionalFormatting();
  assert.strictEqual(s.conditionalFormattings.length, 0);
});

console.log('Data validation');
test('lists, numbers, dates, times, text length, custom and messages', () => {
  const wb = Excel.create();
  const s = wb.sheet('Tasks');
  wb.createSheet('Lists').appendRow(['Open']);
  s.addDataValidation('A2:A100', {
    type: 'list',
    items: ['Open', 'In progress', 'Done'],
    prompt: { title: 'Status', message: 'Pick a status' },
    error: { title: 'Invalid', message: 'Choose from the list', style: 'warning' },
  });
  s.addDataValidation('B2:B100', { type: 'listFromRange', range: 'A1:A20', sheet: 'Lists' });
  s.addDataValidation('C2:C100', { type: 'wholeNumber', operator: 'between', value: 1, value2: 10 });
  s.addDataValidation('D2:D100', { type: 'decimal', operator: 'greaterThan', value: 0.5 });
  s.addDataValidation('E2:E100', { type: 'date', operator: 'greaterThanOrEqual', value: new Date(2026, 0, 1) });
  s.addDataValidation('F2:F100', { type: 'time', operator: 'lessThan', value: '18:00' });
  s.addDataValidation('G2:G100', { type: 'textLength', operator: 'lessThanOrEqual', value: 50 });
  s.addDataValidation('H2:H100', { type: 'custom', formula: 'H2>C2' });
  assert.strictEqual(s.cell('A5').validates('Done'), true);
  assert.strictEqual(s.cell('A5').validates('Nope'), false);
  assert.strictEqual(s.cell('C5').validates(11), false);
  const back = reload(wb).sheet('Tasks');
  assert.deepStrictEqual(back.getDataValidation('A50').items, ['Open', 'In progress', 'Done']);
  assert.strictEqual(back.getDataValidation('A50').error.style, 'warning');
  assert.strictEqual(back.getDataValidation('C2').formula2, '10');
  assert.strictEqual(Object.keys(back.dataValidations).length, 8);
  s.removeDataValidation('A50:A100');
  assert.strictEqual(s.getDataValidation('A60'), null);
  s.clearDataValidations();
  assert.deepStrictEqual(s.dataValidations, {});
});

console.log('Page setup & printing');
test('page setup, margins, print options and header/footer survive save', () => {
  const wb = Excel.create();
  const s = wb.sheet('Report');
  s.cell('A1').value = 'Report';
  s.setPageSetup({ orientation: 'landscape', paperSize: 'a4', fitToWidth: 1, fitToHeight: 0, copies: 2 });
  s.setPageMargins({ left: 2, right: 2, unit: 'cm' });
  s.setPrintOptions({ gridLines: true, headings: true, horizontalCentered: true });
  s.setHeaderFooter({ header: '&CQuarterly report', footer: '&RPage &P of &N' });
  const back = reload(wb).sheet('Report');
  const ps = back.pageSetup;
  assert.deepStrictEqual([ps.orientation, ps.paperSize, ps.fitToWidth, ps.fitToHeight, ps.copies], [
    'landscape',
    'A4',
    1,
    0,
    2,
  ]);
  assert.ok(Math.abs(back.pageMargins.left - 2 / 2.54) < 1e-6);
  assert.deepStrictEqual([back.printOptions.gridLines, back.printOptions.headings], [true, true]);
  assert.strictEqual(back.headerFooter.oddFooter, '&RPage &P of &N');
  s.setPageMargins('narrow');
  assert.strictEqual(s.pageMargins.left, 0.25);
  s.clearPageSetup();
  assert.strictEqual(s.pageSetup, null);
});

console.log('Tables');
test('styles, totals row, read, update, append and remove', () => {
  const wb = Excel.create();
  const s = salesSheet(wb);
  s.appendRow([]);
  s.addTable('A1:C5', 'Sales_T', [
    { name: 'Region', totalsLabel: 'Total' },
    { name: 'Units', totalsFunction: 'sum' },
    { name: 'Price', totalsFunction: 'average' },
  ], { style: 'medium9', showTotalsRow: true });
  let t = reload(wb).sheet('Sales').getTable('Sales_T');
  assert.deepStrictEqual([t.style, t.showTotalsRow, t.columns[1].totalsFunction], [
    'TableStyleMedium9',
    true,
    'sum',
  ]);
  assert.strictEqual(s.tableRowsAsMaps('Sales_T')[0].Region, 'North');
  s.updateTable('Sales_T', { style: 'light1', showColumnStripes: true });
  assert.strictEqual(s.getTable('Sales_T').style, 'TableStyleLight1');
  const t2 = wb.sheet('T2').addTable('A1:B2', 'Small', ['X', 'Y']);
  assert.strictEqual(t2.name, 'Small');
  wb.sheet('T2').appendTableRow('Small', [1, 2]);
  assert.strictEqual(wb.sheet('T2').getTable('Small').ref, 'A1:B3');
  wb.sheet('T2').removeTable('Small');
  assert.strictEqual(wb.sheet('T2').tables.length, 0);
});

console.log('Charts');
test('all 11 chart types with data labels and series styles', () => {
  const wb = Excel.create();
  const s = wb.sheet('Data');
  s.appendRow(['Month', 'Open', 'High', 'Low', 'Close', 'Size']);
  [['Jan', 10, 15, 8, 12, 3], ['Feb', 12, 18, 11, 17, 5], ['Mar', 17, 19, 13, 14, 2]].forEach((r) =>
    s.appendRow(r)
  );
  const series = (col, extra = {}) => ({
    name: col,
    categoriesRange: 'Data!$A$2:$A$4',
    valuesRange: `Data!$${col}$2:$${col}$4`,
    ...extra,
  });
  const types = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'ofPie', 'scatter', 'radar'];
  types.forEach((type, i) =>
    s.addChart({
      type,
      title: type,
      dataLabels: { value: true, percentage: type === 'pie' },
      series: [
        series('B', { style: { fillColor: '#2E86AB', fillType: 'transparent', fillAlpha: 40, borderColor: '#1A5276' } }),
      ],
      anchor: { column: 7, row: i * 16, width: 8, height: 15 },
    })
  );
  s.addChart({
    type: 'bubble',
    series: [series('B', { bubbleSizeRange: 'Data!$F$2:$F$4' })],
    anchor: { column: 16, row: 0 },
  });
  s.addChart({
    type: 'stock',
    series: ['B', 'C', 'D', 'E'].map((c) => series(c)),
    anchor: { column: 16, row: 16 },
  });
  assert.strictEqual(s.chartCount, 11);
  const bytes = wb.encode();
  assert.strictEqual(Excel.read(bytes).sheet('Data').cell('B2').value, 10);
});

console.log('Pivot tables');
test('pivot table with rows, columns and several aggregations', () => {
  const wb = Excel.create();
  salesSheet(wb, 'Data');
  const report = wb.sheet('Pivot');
  report.addPivotTable({
    name: 'PivotTable1',
    sourceSheet: 'Data',
    sourceRange: 'A1:C4',
    targetCell: 'A3',
    rows: ['Region'],
    values: [
      { field: 'Units', function: 'sum', customName: 'Total units' },
      { field: 'Price', function: 'average' },
    ],
  });
  assert.strictEqual(report.pivotTableCount, 1);
  assert.ok(reload(wb).sheets.includes('Pivot'));
});

console.log('Sheet options');
test('grouping: collapsed groups, expand, levels, settings and clear', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 1;
  s.groupRows(1, 8);
  s.groupRows(2, 4, { collapsed: true });
  s.groupColumns(1, 3, { collapsed: true });
  assert.strictEqual(s.getRowOutlineLevel(3), 2);
  assert.deepStrictEqual(s.rowGroups.map((g) => [g.start, g.end, g.level]), [[1, 8, 1], [2, 4, 2]]);
  s.expandColumnGroup(1, 3);
  assert.strictEqual(s.columnGroups[0].collapsed, false);
  s.outlineSettings = { summaryBelow: false };
  assert.strictEqual(reload(wb).sheet('Sheet1').outlineSettings.summaryBelow, false);
  s.clearGrouping();
  assert.strictEqual(s.rowGroups.length, 0);
});

test('autofilter column criteria', () => {
  const wb = Excel.create();
  const s = salesSheet(wb);
  s.setAutoFilter('A1:D4');
  s.addFilterColumn({ column: 0, values: ['North', 'East'] });
  s.addFilterColumn({ column: 1, custom: [{ operator: 'greaterThan', value: 5 }] });
  const f = reload(wb).sheet('Sales').autoFilter;
  assert.strictEqual(f.ref, 'A1:D4');
  assert.deepStrictEqual(f.columns[0].values, ['North', 'East']);
  assert.deepStrictEqual(f.columns[1].custom, [{ operator: 'greaterThan', value: '5' }]);
});

test('protection options, tab theme color, image offsets', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.protect('secret', { selectLockedCells: false, insertHyperlinks: true });
  assert.strictEqual(s.isProtected, true);
  s.setTabColorTheme(4, 0.4);
  s.addImage(PNG, 'png', 1, 1, 20, 20, { colOffset: 5, rowOffset: 5 });
  assert.ok(reload(wb).sheet('Sheet1').isProtected);
});

test('linked sheets share their content', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 'shared';
  wb.linkSheet('Mirror', 'Sheet1');
  wb.sheet('Mirror').cell('A2').value = 'from mirror';
  assert.strictEqual(wb.sheet('Sheet1').cell('A2').value, 'from mirror');
  wb.unlinkSheet('Mirror');
  wb.sheet('Mirror').cell('A3').value = 'only mirror';
  assert.strictEqual(wb.sheet('Sheet1').cell('A3').value, null);
});

console.log('Full workbook');
test('a workbook with every feature encodes twice identically and reads back', () => {
  const wb = Excel.create();
  const s = salesSheet(wb);
  s.addConditionalFormatting('B2:B4', { type: 'cellIs', operator: 'greaterThan', value: 5, style: { bold: true } });
  s.addDataValidation('A2:A10', { type: 'list', items: ['North', 'South', 'East'] });
  s.setPageSetup({ orientation: 'landscape' });
  s.addTable('A1:D4', 'Full', null, { style: 'medium2' });
  // Pie/doughnut slice colors are shuffled on purpose, so use a column chart here.
  s.addChart({ type: 'column', series: [{ categoriesRange: 'Sales!$A$2:$A$4', valuesRange: 'Sales!$B$2:$B$4' }] });
  s.cell('F1').setHyperlink({ url: 'https://dart.dev' });
  s.groupRows(1, 2);
  wb.sheet('Pivot').addPivotTable({
    sourceSheet: 'Sales',
    sourceRange: 'A1:B4',
    targetCell: 'A1',
    rows: ['Region'],
    values: [{ field: 'Units' }],
  });
  assert.deepStrictEqual(Buffer.from(wb.encode()), Buffer.from(wb.encode()));
  assert.ok(reload(wb).sheets.includes('Pivot'));
  fs.writeFileSync(path.join(__dirname, 'output_features_test.xlsx'), wb.encode());
});

if (failed) {
  console.log(`\n${failed} feature test(s) failed`);
  process.exit(1);
}
fs.unlinkSync(path.join(__dirname, 'output_features_test.xlsx'));
console.log('\n🎉 All feature tests passed');
