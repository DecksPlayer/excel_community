const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { Excel } = require('./index.js');

console.log('Testing full feature set of excel-community npm package...');

// 1. Create workbook & set cell values
const wb = Excel.create();
const sheet = wb.sheet('Sheet1');

sheet.cell('A1').value = 'Title';
sheet.cell('B1').value = 100;
sheet.cell('C1').value = 45.67;
sheet.cell('D1').value = true;
sheet.cell('E1').formula = '=SUM(B1:C1)';

assert.strictEqual(sheet.cell('A1').value, 'Title');
assert.strictEqual(sheet.cell('B1').value, 100);
assert.strictEqual(sheet.cell('C1').value, 45.67);
assert.strictEqual(sheet.cell('D1').value, true);
assert.strictEqual(sheet.cell('E1').formula, '=SUM(B1:C1)');

// 2. Full Cell Styling: fonts, colors, alignments, borders, number format
sheet.cell('A1').setStyle({
  bold: true,
  italic: true,
  fontSize: 14,
  fontFamily: 'Arial',
  fontColor: '#FFFFFF',
  backgroundColor: '#1F4E78',
  horizontalAlign: 'center',
  verticalAlign: 'center',
  border: { style: 'thin', color: '#002060' },
  numberFormat: '$#,##0.00',
});

const style = sheet.cell('A1').style;
assert.ok(style);
assert.strictEqual(style.bold, true);
assert.strictEqual(style.italic, true);
assert.strictEqual(style.fontSize, 14);
assert.strictEqual(style.horizontalAlign, 'Center');

// 3. Hyperlinks
sheet.cell('A1').setHyperlink('https://google.com', 'Google Search', 'Visit Google');
const link = sheet.cell('A1').getHyperlink();
assert.ok(link);
assert.strictEqual(link.url, 'https://google.com');
sheet.cell('A1').removeHyperlink();
assert.strictEqual(sheet.cell('A1').getHyperlink(), null);

// 4. Batch append row and 2D rows traversal
sheet.appendRow(['Item 1', 25.5, false]);
assert.ok(sheet.rows.length >= 2);
const firstCell = sheet.rows[0][0];
assert.strictEqual(firstCell.value, 'Visit Google');

// 5. Column widths, row heights, auto-fit, default dimensions
sheet.setColumnWidth(0, 25);
assert.strictEqual(sheet.getColumnWidth(0), 25);
sheet.setRowHeight(0, 30);
assert.strictEqual(sheet.getRowHeight(0), 30);
sheet.setDefaultColumnWidth(15);
sheet.setDefaultRowHeight(20);
sheet.setColumnAutoFit(1);

// 6. Merging and unmerging
sheet.merge('F1', 'G1');
assert.ok(sheet.spannedItems.includes('F1:G1'));
sheet.unmerge('G1');
assert.ok(!sheet.spannedItems.includes('F1:G1'));
sheet.merge('F1', 'G1');
sheet.unmerge('F1:G1');
assert.ok(!sheet.spannedItems.includes('F1:G1'));

// 7. AutoFilter & Tab color & RTL
sheet.setAutoFilter('A1:E2');
assert.strictEqual(sheet.hasAutoFilter, true);
sheet.clearAutoFilter();
assert.strictEqual(sheet.hasAutoFilter, false);

sheet.tabColor = '#FF0000';
assert.ok(sheet.tabColor);

sheet.rightToLeft = true;
assert.strictEqual(sheet.rightToLeft, true);
sheet.rightToLeft = false;

// 8. Protection
sheet.protect('password123', { formatCells: false });
sheet.unprotect();

// 9. Grouping / Outline
sheet.groupRows(1, 2);
sheet.ungroupRows(1, 2);
sheet.groupColumns(0, 1);
sheet.ungroupColumns(0, 1);

// 10. Charts (Column, Bar, Line, Pie, Area, Scatter, Radar)
sheet.addChart(JSON.stringify({
  type: 'column',
  title: 'Sales Overview',
  series: [
    { name: '2026', categoriesRange: 'Sheet1!$A$1:$A$2', valuesRange: 'Sheet1!$B$1:$B$2', colorHex: '4472C4' }
  ],
  anchor: { fromCol: 5, fromRow: 2, toCol: 13, toRow: 16 }
}));

// 11. Tables
sheet.addTable('A10:B12', 'MyTable');

// 12. Workbook sheet management
assert.ok(wb.sheets.includes('Sheet1'));
wb.copySheet('Sheet1', 'Sheet1_Copy');
assert.ok(wb.sheets.includes('Sheet1_Copy'));
wb.deleteSheet('Sheet1_Copy');
assert.ok(!wb.sheets.includes('Sheet1_Copy'));

// 13. End-to-end Save & Read back from file
const testFile = path.join(__dirname, 'output_full_test.xlsx');
wb.save(testFile);
assert.ok(fs.existsSync(testFile));

const loaded = Excel.fromFile(testFile);
const loadedSheet = loaded.sheet('Sheet1');
assert.strictEqual(loadedSheet.cell('A1').value, 'Visit Google');
assert.strictEqual(loadedSheet.cell('B1').value, 100);
assert.strictEqual(loadedSheet.cell('C1').value, 45.67);
assert.strictEqual(loadedSheet.cell('D1').value, true);

// Cleanup
fs.unlinkSync(testFile);

// 14. Dates, appendRow column positions, config objects, plain arrays
const wb2 = Excel.create();
const s2 = wb2.sheet('Sheet1');
assert.deepStrictEqual(wb2.sheets, ['Sheet1']);
s2.appendRow(['a', new Date(2024, 0, 15), {}, 'd', new Date(2024, 0, 15, 10, 30)]);
assert.deepStrictEqual(
  s2.rows[0].map((c) => c && c.type),
  ['string', 'date', 'null', 'string', 'datetime']
);
s2.cell('A3').value = new Date(2026, 9, 4);
s2.addChart({ type: 'line', series: [{ categoriesRange: 'Sheet1!$A$1:$A$2', valuesRange: 'Sheet1!$B$1:$B$2' }] });
s2.addTable('A10:B12', 'ObjTable', ['X', 'Y']);
const s2back = Excel.read(wb2.encode()).sheet('Sheet1');
assert.strictEqual(s2back.cell('B1').value, '2024-01-15');
assert.strictEqual(s2back.cell('E1').value, '2024-01-15T10:30:00.000');
assert.strictEqual(s2back.cell('A3').value, '2026-10-04');

// 15. Encoding twice gives the same file (no duplicated drawings or styles)
assert.deepStrictEqual(Buffer.from(wb2.encode()), Buffer.from(wb2.encode()));

// 16. No browser globals leaked into Node
assert.strictEqual(typeof window, 'undefined');
assert.strictEqual(typeof self, 'undefined');

console.log('🎉 ALL functional tests for excel_community passed successfully!');
