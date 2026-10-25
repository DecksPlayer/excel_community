/* global ExcelCommunity */
const { Excel } = ExcelCommunity;

const $ = (sel) => document.querySelector(sel);

// ==================== ASSERTIONS ====================

function same(actual, expected, label = 'value') {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error(`${label}: expected ${e}, got ${a}`);
}

function ok(condition, message = 'assertion failed') {
  if (!condition) throw new Error(message);
}

function throws(fn, label) {
  try {
    fn();
  } catch (e) {
    return;
  }
  throw new Error(`${label} should throw`);
}

/** Saves and reads the workbook back, as Excel would. */
function roundTrip(wb) {
  return Excel.read(wb.encode());
}

async function pngBytes(color = '#2E75B6', size = 64) {
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = size;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, size, size);
  ctx.fillStyle = '#FFFFFF';
  ctx.font = `bold ${size / 2}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('EC', size / 2, size / 2);
  // Some browser extensions block canvas export, so `toBlob` may never call back.
  const blob = await Promise.race([
    new Promise((resolve) => canvas.toBlob(resolve, 'image/png')),
    new Promise((resolve) => setTimeout(() => resolve(null), 2000)),
  ]);
  if (blob) return new Uint8Array(await blob.arrayBuffer());
  // Canvas export unavailable: use a built-in 1x1 PNG.
  return Uint8Array.from(
    atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
    (c) => c.charCodeAt(0)
  );
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
let fallbackUrl = null;

/**
 * Downloads the workbook with `save()` and reports the result on the page, with
 * a direct link as a fallback in case the browser blocked or cancelled the download.
 */
async function downloadWorkbook(wb, fileName, statusEl) {
  await showStep(statusEl, 'Encoding…');
  const bytes = wb.encode();
  await showStep(statusEl, 'Saving…');
  wb.save(fileName);
  if (fallbackUrl) URL.revokeObjectURL(fallbackUrl);
  fallbackUrl = URL.createObjectURL(new Blob([bytes], { type: XLSX_MIME }));
  const link = document.createElement('a');
  link.href = fallbackUrl;
  link.download = fileName;
  link.textContent = 'download it from this link';
  statusEl.className = 'status';
  statusEl.replaceChildren(
    `Download requested: ${fileName} (${(bytes.length / 1024).toFixed(1)} KB). Didn't get it? `,
    link,
    '.'
  );
}

/** Shows [text] and lets the browser paint it before the next (blocking) step. */
function showStep(statusEl, text) {
  statusEl.className = 'status';
  statusEl.textContent = text;
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

function showError(statusEl, e) {
  console.error(e);
  statusEl.className = 'status error';
  statusEl.textContent = `Error: ${(e && e.message) || e}`;
}

// ==================== TEST SUITE ====================

const suite = [];
const test = (group, name, fn) => suite.push({ group, name, fn });

// ---- Workbook ----
test('Workbook', 'create() starts with Sheet1', () => {
  same(Excel.create().sheets, ['Sheet1'], 'sheets');
});

test('Workbook', 'sheet() on a new workbook renames the empty Sheet1', () => {
  const wb = Excel.create();
  wb.sheet('Sales');
  same(wb.sheets, ['Sales'], 'sheets');
});

test('Workbook', 'createSheet / renameSheet / copySheet / deleteSheet', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 'keep';
  wb.createSheet('Inventory');
  same(wb.sheets, ['Sheet1', 'Inventory'], 'after create');
  wb.renameSheet('Sheet1', 'Summary');
  same(wb.sheets, ['Inventory', 'Summary'], 'after rename');
  wb.copySheet('Summary', 'Backup');
  same(wb.sheet('Backup').cell('A1').value, 'keep', 'copied value');
  wb.deleteSheet('Backup');
  ok(!wb.sheets.includes('Backup'), 'Backup should be deleted');
});

test('Workbook', 'defaultSheet', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 1;
  wb.createSheet('Second');
  wb.defaultSheet = 'Second';
  same(wb.defaultSheet, 'Second', 'defaultSheet');
  same(roundTrip(wb).defaultSheet, 'Second', 'defaultSheet after reload');
});

test('Workbook', 'encode() returns a ZIP (.xlsx) Uint8Array', () => {
  const bytes = Excel.create().encode();
  ok(bytes instanceof Uint8Array, 'not a Uint8Array');
  same([bytes[0], bytes[1]], [0x50, 0x4b], 'ZIP signature "PK"');
});

test('Workbook', 'toBuffer() returns bytes', () => {
  const bytes = Excel.create().toBuffer();
  ok(bytes instanceof Uint8Array && bytes.length > 0, 'empty buffer');
});

test('Workbook', 'read() accepts Uint8Array, ArrayBuffer, typed views and Base64', () => {
  const wb = Excel.create();
  wb.sheet('Data').cell('A1').value = 'hello';
  const bytes = wb.encode();
  same(Excel.read(bytes).sheet('Data').cell('A1').value, 'hello', 'Uint8Array');
  same(Excel.read(bytes.slice().buffer).sheets, ['Data'], 'ArrayBuffer');
  const padded = new Uint8Array(bytes.length + 8);
  padded.set(bytes, 4);
  same(Excel.read(padded.subarray(4, 4 + bytes.length)).sheets, ['Data'], 'offset view');
  same(Excel.read(wb.encodeBase64()).sheets, ['Data'], 'Base64');
});

test('Workbook', 'read() rejects invalid data', () => {
  throws(() => Excel.read(new Uint8Array([1, 2, 3])), 'invalid bytes');
});

// ---- Cells ----
test('Cells', 'string, int, double and bool values', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').value = 'text';
  s.cell('A2').value = 42;
  s.cell('A3').value = 3.5;
  s.cell('A4').value = true;
  same(['A1', 'A2', 'A3', 'A4'].map((c) => s.cell(c).type), ['string', 'int', 'double', 'bool'], 'types');
  same(['A1', 'A2', 'A3', 'A4'].map((c) => s.cell(c).value), ['text', 42, 3.5, true], 'values');
});

test('Cells', 'Date objects become date / datetime cells', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = new Date(2026, 9, 4);
  s.cell('A2').value = new Date(2026, 9, 4, 14, 30);
  const back = roundTrip(wb).sheet('Sheet1');
  same([back.cell('A1').type, back.cell('A1').value], ['date', '2026-10-04'], 'date');
  same([back.cell('A2').type, back.cell('A2').value], ['datetime', '2026-10-04T14:30:00.000'], 'datetime');
});

test('Cells', 'null clears a cell', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').value = 'x';
  s.cell('A1').value = null;
  same([s.cell('A1').value, s.cell('A1').type], [null, 'null'], 'cleared cell');
});

test('Cells', 'formulas (property, setFormula and "=" strings)', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 2;
  s.cell('A2').value = 3;
  s.cell('A3').formula = '=SUM(A1:A2)';
  s.cell('A4').setFormula('=A1*A2');
  s.cell('A5').value = '=A1+A2';
  const back = roundTrip(wb).sheet('Sheet1');
  same(back.cell('A3').type, 'formula', 'type');
  ok(/SUM\(A1:A2\)/.test(back.cell('A3').formula), 'A3 formula lost');
  ok(/A1\*A2/.test(back.cell('A4').formula), 'A4 formula lost');
  ok(/A1\+A2/.test(back.cell('A5').formula), 'A5 formula lost');
});

test('Cells', 'cellId, row and col', () => {
  const c = Excel.create().sheet('Sheet1').cell('AB12');
  same([c.cellId, c.row, c.col], ['AB12', 11, 27], 'coordinates');
});

test('Cells', 'comments survive save', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('B2').comment = 'Reviewed';
  same(roundTrip(wb).sheet('Sheet1').cell('B2').comment, 'Reviewed', 'comment');
});

test('Cells', 'displayText applies the number format', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').value = 1250.5;
  s.cell('A1').setStyle({ numberFormat: '$#,##0.00' });
  same(s.cell('A1').displayText, '$1,250.50', 'displayText');
});

test('Cells', 'hyperlinks: set, get and remove', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').setHyperlink('https://pub.dev', 'Open pub.dev', 'pub.dev');
  same(s.cell('A1').getHyperlink().url, 'https://pub.dev', 'url');
  same(s.cell('A1').value, 'pub.dev', 'display text');
  same(roundTrip(wb).sheet('Sheet1').cell('A1').getHyperlink().url, 'https://pub.dev', 'url after reload');
  s.cell('A1').removeHyperlink();
  same(s.cell('A1').getHyperlink(), null, 'removed');
});

// ---- Styles ----
test('Styles', 'font, colors and alignment survive save', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 'Styled';
  wb.sheet('Sheet1').cell('A1').setStyle({
    bold: true,
    italic: true,
    underline: 'double',
    strikethrough: true,
    fontSize: 16,
    fontFamily: 'Arial',
    fontColor: '#FFFFFF',
    backgroundColor: '#1F4E78',
    horizontalAlign: 'center',
    verticalAlign: 'top',
    rotation: 45,
  });
  const st = roundTrip(wb).sheet('Sheet1').cell('A1').style;
  same(
    [st.bold, st.italic, st.strikethrough, st.underline, st.fontSize, st.fontFamily],
    [true, true, true, 'double', 16, 'Arial'],
    'font'
  );
  same([st.fontColor, st.backgroundColor], ['#FFFFFF', '#1F4E78'], 'colors');
  same([st.horizontalAlign, st.verticalAlign, st.rotation], ['center', 'top', 45], 'alignment');
});

test('Styles', 'borders (all sides and per side)', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').setStyle({ border: { style: 'thin', color: '#002060' } });
  s.cell('B2').setStyle({ bottomBorder: { style: 'double' }, leftBorder: { style: 'dashed' } });
  ok(s.cell('A1').style, 'missing style');
});

test('Styles', 'built-in and custom number formats', () => {
  const s = Excel.create().sheet('Sheet1');
  s.cell('A1').value = 0.256;
  s.cell('A1').setStyle({ numberFormat: 10 });
  same(s.cell('A1').displayText, '25.60%', 'format 10');
  s.cell('A2').value = 1234567;
  s.cell('A2').setStyle({ numberFormat: '#,##0' });
  same(s.cell('A2').displayText, '1,234,567', 'custom');
});

// ---- Sheets ----
test('Sheets', 'appendRow keeps column positions and rows returns the grid', () => {
  const s = Excel.create().sheet('Sheet1');
  s.appendRow(['ID', 'Product', 'Qty', 'Price', 'Total']);
  s.appendRow([1, 'Keyboard', 2, 85.5, '=C2*D2']);
  s.appendRow([2, null, 5, undefined, new Date(2026, 0, 1)]);
  same([s.maxRows, s.maxColumns], [3, 5], 'dimensions');
  same(s.rows[1].map((c) => c && c.value), [1, 'Keyboard', 2, 85.5, '=C2*D2'], 'row 2');
  same(s.rows[2].map((c) => c && c.type), ['int', 'null', 'int', 'null', 'date'], 'row 3 types');
});

test('Sheets', 'column widths and row heights', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.setColumnWidth(0, 30);
  s.setRowHeight(0, 40);
  s.setDefaultColumnWidth(14);
  s.setDefaultRowHeight(18);
  s.cell('B1').value = 'A long piece of text for auto-fit';
  s.setColumnAutoFit(1);
  const back = roundTrip(wb).sheet('Sheet1');
  same([back.getColumnWidth(0), back.getRowHeight(0)], [30, 40], 'sizes after reload');
});

test('Sheets', 'freeze panes', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 'Header';
  s.frozenRows = 1;
  s.frozenColumns = 2;
  const back = roundTrip(wb).sheet('Sheet1');
  same([back.frozenRows, back.frozenColumns], [1, 2], 'frozen after reload');
  s.frozenRows = null;
  same(s.frozenRows, null, 'unfrozen');
});

test('Sheets', 'hidden rows and columns', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 1;
  s.setColumnHidden(1, true);
  s.setRowHidden(4, true);
  const back = roundTrip(wb).sheet('Sheet1');
  same([back.isColumnHidden(1), back.isRowHidden(4), back.isRowHidden(3)], [true, true, false], 'hidden');
});

test('Sheets', 'merge and unmerge', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 'Banner';
  s.merge('A1', 'E1');
  same(s.spannedItems, ['A1:E1'], 'spannedItems');
  same(roundTrip(wb).sheet('Sheet1').spannedItems, ['A1:E1'], 'after reload');
  s.unmerge('A1');
  same(s.spannedItems, [], 'after unmerge');
});

test('Sheets', 'tab color and right-to-left', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.tabColor = '#FF0000';
  s.rightToLeft = true;
  const back = roundTrip(wb).sheet('Sheet1');
  same([back.tabColor, back.rightToLeft], ['#FF0000', true], 'after reload');
  s.tabColor = null;
  same(s.tabColor, null, 'cleared');
});

test('Sheets', 'auto filter', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.appendRow(['Name', 'Score']);
  s.appendRow(['Ana', 9]);
  s.setAutoFilter('A1:B2');
  ok(roundTrip(wb).sheet('Sheet1').hasAutoFilter, 'filter lost on reload');
  s.clearAutoFilter();
  same(s.hasAutoFilter, false, 'cleared');
});

test('Sheets', 'protection with password and options', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.cell('A1').value = 'locked';
  s.protect('secret', { formatCells: false, sort: true });
  same(roundTrip(wb).sheet('Sheet1').cell('A1').value, 'locked', 'value after reload');
  s.unprotect();
});

test('Sheets', 'row and column grouping', () => {
  const s = Excel.create().sheet('Sheet1');
  s.groupRows(1, 3);
  s.groupColumns(1, 2);
  s.ungroupRows(1, 3);
  s.ungroupColumns(1, 2);
});

test('Sheets', 'images (PNG generated with canvas)', async () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').addImage(await pngBytes(), 'png', 1, 1, 64, 64);
  ok(roundTrip(wb).sheets.includes('Sheet1'), 'reload failed');
});

const chartTypes = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'ofPie', 'scatter', 'radar', 'bubble', 'stock'];
test('Charts', `all ${chartTypes.length} chart types with data labels and series styles`, () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.appendRow(['Q', 'Open', 'High', 'Low', 'Close']);
  [['Q1', 10, 15, 8, 12], ['Q2', 12, 18, 11, 17], ['Q3', 17, 19, 13, 14]].forEach((r) => s.appendRow(r));
  const series = (col, extra) => ({
    name: col,
    categoriesRange: 'Sheet1!$A$2:$A$4',
    valuesRange: `Sheet1!$${col}$2:$${col}$4`,
    ...extra,
  });
  chartTypes.forEach((type, i) =>
    s.addChart({
      type,
      title: type,
      dataLabels: { value: true },
      series:
        type === 'stock'
          ? ['B', 'C', 'D', 'E'].map((c) => series(c))
          : [series('B', { bubbleSizeRange: 'Sheet1!$C$2:$C$4', style: { fillColor: '#2E86AB', fillType: 'transparent', fillAlpha: 40 } })],
      anchor: { column: 6, row: i * 16, width: 8, height: 15 },
    })
  );
  same(s.chartCount, chartTypes.length, 'chartCount');
  ok(roundTrip(wb).sheet('Sheet1').cell('B2').value === 10, 'data lost');
});

test('Charts', 'config as JSON string, stacked grouping and custom colors', () => {
  const s = Excel.create().sheet('Sheet1');
  s.addChart(
    JSON.stringify({
      type: 'column',
      grouping: 'stacked',
      showLegend: false,
      series: [
        { name: 'A', categoriesRange: 'Sheet1!$A$1:$A$2', valuesRange: 'Sheet1!$B$1:$B$2', colorHex: 'ED7D31' },
      ],
    })
  );
});

test('Tables', 'tables with array and JSON columns', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  s.appendRow(['Name', 'Score']);
  s.appendRow(['Ana', 9]);
  s.addTable('A1:B2', 'Scores', ['Name', 'Score']);
  s.addTable('D1:E2', 'Other', JSON.stringify(['X', 'Y']));
  ok(roundTrip(wb).sheet('Sheet1').cell('A2').value === 'Ana', 'data lost');
});

test('Tables', 'invalid table names are rejected', () => {
  throws(() => Excel.create().sheet('Sheet1').addTable('A1:B2', 'T2'), 'name "T2"');
});

/** A sheet with a header row and three data rows. */
function salesData(wb, name = 'Sales') {
  const s = wb.sheet(name);
  s.appendRow(['Region', 'Units', 'Price']);
  [['North', 10, 2.5], ['South', 7, 3], ['East', 4, 1.5]].forEach((r) => s.appendRow(r));
  return s;
}

test('Tables', 'style, totals row, update, append rows and remove', () => {
  const wb = Excel.create();
  const s = salesData(wb);
  s.appendRow([]);
  s.addTable('A1:C5', 'Sales_T', [
    { name: 'Region', totalsLabel: 'Total' },
    { name: 'Units', totalsFunction: 'sum' },
    { name: 'Price', totalsFunction: 'average' },
  ], { style: 'medium9', showTotalsRow: true });
  const t = roundTrip(wb).sheet('Sales').getTable('Sales_T');
  same([t.style, t.showTotalsRow, t.columns[1].totalsFunction], ['TableStyleMedium9', true, 'sum'], 'table');
  same(s.tableRowsAsMaps('Sales_T')[0].Region, 'North', 'tableRowsAsMaps');
  s.updateTable('Sales_T', { style: 'light1' });
  same(s.getTable('Sales_T').style, 'TableStyleLight1', 'updated style');
  s.removeTable('Sales_T');
  same(s.tables.length, 0, 'removed');
});

test('Cells', 'setStyle keeps earlier options; resetStyle clears them', () => {
  const c = Excel.create().sheet('Sheet1').cell('A1');
  c.value = 1234.5;
  c.setStyle({ numberFormat: '$#,##0.00' });
  c.setStyle({ bold: true });
  same([c.style.bold, c.displayText], [true, '$1,234.50'], 'merged style');
  c.resetStyle();
  same(c.style.bold, false, 'reset');
});

test('Cells', 'times, Date values and every hyperlink type', () => {
  const wb = Excel.create();
  const s = wb.sheet('Sheet1');
  wb.createSheet('Other');
  s.cell('A1').setTime('14:30');
  s.cell('A2').value = new Date(2026, 9, 4);
  s.cell('B1').setHyperlink({ email: 'sales@example.com', subject: 'Q3' }, { text: 'Mail' });
  s.cell('B2').setHyperlink({ sheet: 'Other', cell: 'C3' }, { text: 'Go' });
  s.setHyperlinkRange('C1:C3', { location: 'Sheet1!A1' });
  const back = roundTrip(wb).sheet('Sheet1');
  same(back.cell('A1').value, '14:30:00', 'time');
  ok(back.cell('A2').dateValue instanceof Date, 'dateValue should be a Date');
  same(back.cell('B1').getHyperlink().url, 'mailto:sales@example.com?subject=Q3', 'email');
  same(back.cell('B2').getHyperlink().location, 'Other!C3', 'cell link');
  ok(Object.keys(back.hyperlinks).includes('C1:C3'), 'range link');
});

test('Rows & columns', 'insert, remove, clear, merge value, ranges and find & replace', () => {
  const s = salesData(Excel.create());
  s.insertRow(1);
  same(s.cell('A3').value, 'North', 'insertRow');
  s.removeRow(1);
  s.insertColumn(0);
  same(s.cell('B1').value, 'Region', 'insertColumn');
  s.removeColumn(0);
  same(s.rangeValues('A2:B3'), [['North', 10], ['South', 7]], 'rangeValues');
  same(s.findAndReplace(/north|south/i, 'Center'), 2, 'replacements');
  s.merge('E1', 'G1', 'Banner');
  same(s.cell('E1').value, 'Banner', 'merge value');
  same(s.clearRow(3), true, 'clearRow');
});

test('Export & import', 'rowsAsMaps, toJson, toCsv and appendRowsFromMaps', () => {
  const wb = Excel.create();
  const s = salesData(wb);
  same(s.rowsAsMaps()[1], { Region: 'South', Units: 7, Price: 3 }, 'rowsAsMaps');
  same(JSON.parse(s.toJson())[2].Region, 'East', 'toJson');
  ok(s.toCsv().startsWith('Region,Units,Price'), 'toCsv');
  same(Object.keys(wb.toMaps()), ['Sales'], 'workbook toMaps');
  const imported = wb.sheet('Imported');
  imported.appendRowsFromMaps([{ Name: 'Eva', Age: 40 }, { Age: 31, Name: 'Ana' }]);
  same(imported.rowsAsValues(), [['Name', 'Age'], ['Eva', 40], ['Ana', 31]], 'appendRowsFromMaps');
});

test('Conditional formatting', 'cellIs, between, expression, text, duplicates', () => {
  const wb = Excel.create();
  const s = salesData(wb);
  s.addConditionalFormatting('B2:B4', { type: 'cellIs', operator: 'greaterThan', value: 5, style: { backgroundColor: '#C6EFCE', bold: true } });
  s.addConditionalFormatting('C2:C4', [
    { type: 'cellIs', operator: 'between', value: 1, value2: 2, style: { fontColor: '#9C0006' } },
    { type: 'expression', formula: '$C2>2.5', style: { italic: true }, priority: 2 },
  ]);
  s.addConditionalFormatting('A2:A4', { type: 'containsText', text: 'th', style: { underline: 'single' } });
  s.addConditionalFormatting('A2:A4', { type: 'duplicateValues', style: { backgroundColor: '#FFC7CE' } });
  const back = roundTrip(wb).sheet('Sales').conditionalFormattings;
  same(back.length, 4, 'groups');
  same(back[1].rules.map((r) => r.type), ['cellIs', 'expression'], 'rule types');
});

test('Data validation', 'dropdowns, numbers, dates, custom rules and messages', () => {
  const wb = Excel.create();
  const s = wb.sheet('Tasks');
  s.addDataValidation('A2:A50', {
    type: 'list',
    items: ['Open', 'Done'],
    prompt: { title: 'Status', message: 'Pick a status' },
    error: { title: 'Invalid', message: 'Use the list', style: 'warning' },
  });
  s.addDataValidation('B2:B50', { type: 'wholeNumber', operator: 'between', value: 1, value2: 10 });
  s.addDataValidation('C2:C50', { type: 'date', operator: 'greaterThan', value: new Date(2026, 0, 1) });
  s.addDataValidation('D2:D50', { type: 'custom', formula: 'D2>B2' });
  same([s.cell('A5').validates('Done'), s.cell('A5').validates('Nope'), s.cell('B5').validates(11)], [true, false, false], 'validates');
  const back = roundTrip(wb).sheet('Tasks');
  same(back.getDataValidation('A9').items, ['Open', 'Done'], 'list items');
  same(Object.keys(back.dataValidations).length, 4, 'rules');
});

test('Page setup', 'orientation, paper, fit, margins, print options, header & footer', () => {
  const wb = Excel.create();
  const s = wb.sheet('Report');
  s.cell('A1').value = 'Report';
  s.setPageSetup({ orientation: 'landscape', paperSize: 'a4', fitToWidth: 1, fitToHeight: 0 });
  s.setPageMargins('narrow');
  s.setPrintOptions({ gridLines: true, horizontalCentered: true });
  s.setHeaderFooter({ header: '&CSales & costs', footer: '&RPage &P of &N' });
  const back = roundTrip(wb).sheet('Report');
  same([back.pageSetup.orientation, back.pageSetup.paperSize, back.pageSetup.fitToWidth], ['landscape', 'A4', 1], 'page setup');
  same(back.pageMargins.left, 0.25, 'margins');
  same(back.printOptions.gridLines, true, 'gridLines');
  same(back.headerFooter.oddHeader, '&CSales & costs', 'header text');
});

test('Pivot tables', 'rows, columns and several aggregations', () => {
  const wb = Excel.create();
  salesData(wb, 'Data');
  const p = wb.sheet('Pivot');
  p.addPivotTable({
    sourceSheet: 'Data',
    sourceRange: 'A1:C4',
    targetCell: 'A3',
    rows: ['Region'],
    values: [{ field: 'Units', function: 'sum', customName: 'Total units' }, { field: 'Price', function: 'average' }],
  });
  same(p.pivotTableCount, 1, 'pivotTableCount');
  ok(roundTrip(wb).sheets.includes('Pivot'), 'reload');
});

test('Sheets', 'collapsible groups, outline settings, filter criteria, theme tab color', () => {
  const wb = Excel.create();
  const s = salesData(wb);
  s.groupRows(1, 3);
  s.groupRows(2, 3, { collapsed: true });
  same(s.rowGroups.map((g) => g.level), [1, 2], 'group levels');
  s.expandRowGroup(2, 3);
  s.outlineSettings = { summaryBelow: false };
  s.setAutoFilter('A1:C4');
  s.addFilterColumn({ column: 1, custom: [{ operator: 'greaterThan', value: 5 }] });
  s.setTabColorTheme(5, 0.4);
  s.protect('pw', { selectLockedCells: false });
  const back = roundTrip(wb).sheet('Sales');
  same(back.outlineSettings.summaryBelow, false, 'outline settings');
  same(back.autoFilter.columns[0].custom, [{ operator: 'greaterThan', value: '5' }], 'filter');
  same(back.isProtected, true, 'protected');
});

test('Workbook', 'linked sheets share their content', () => {
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 'shared';
  wb.linkSheet('Mirror', 'Sheet1');
  wb.sheet('Mirror').cell('A2').value = 'x';
  same(wb.sheet('Sheet1').cell('A2').value, 'x', 'linked');
  wb.unlinkSheet('Mirror');
});

async function runTests() {
  const list = $('#results');
  list.replaceChildren();
  document.body.dataset.status = 'running';
  let passed = 0;
  const started = performance.now();
  let groupEl = null;
  let currentGroup = null;

  const groups = [...new Set(suite.map((t) => t.group))];
  const ordered = [...suite].sort((a, b) => groups.indexOf(a.group) - groups.indexOf(b.group));
  for (const { group, name, fn } of ordered) {
    if (group !== currentGroup) {
      currentGroup = group;
      groupEl = document.createElement('li');
      groupEl.className = 'group';
      groupEl.innerHTML = `<h3></h3><ul></ul>`;
      groupEl.querySelector('h3').textContent = group;
      list.append(groupEl);
    }
    const item = document.createElement('li');
    const t0 = performance.now();
    let error = null;
    try {
      await fn();
      passed++;
    } catch (e) {
      error = e && e.message ? e.message : String(e);
    }
    item.className = error ? 'fail' : 'pass';
    item.innerHTML = `<span class="mark"></span><span class="name"></span><span class="ms"></span>`;
    item.querySelector('.mark').textContent = error ? '✗' : '✓';
    item.querySelector('.name').textContent = name;
    item.querySelector('.ms').textContent = `${(performance.now() - t0).toFixed(0)} ms`;
    if (error) {
      const pre = document.createElement('pre');
      pre.textContent = error;
      item.append(pre);
    }
    groupEl.querySelector('ul').append(item);
  }

  const failed = suite.length - passed;
  const summary = $('#summary');
  summary.className = failed ? 'summary fail' : 'summary pass';
  summary.textContent =
    `${passed}/${suite.length} passed` +
    (failed ? ` · ${failed} failed` : '') +
    ` · ${(performance.now() - started).toFixed(0)} ms`;
  document.body.dataset.status = failed ? 'failed' : 'passed';
}

// ==================== DEMO WORKBOOK ====================

async function buildDemo(step = async () => {}) {
  await step('Building sheets…');
  const wb = Excel.create();
  const s = wb.sheet('Dashboard');
  const header = {
    bold: true,
    fontColor: '#FFFFFF',
    backgroundColor: '#1F4E78',
    horizontalAlign: 'center',
    border: { style: 'thin', color: '#0F2740' },
  };

  s.cell('A1').value = 'Quarterly Sales Dashboard';
  s.cell('A1').setStyle({ bold: true, fontSize: 18, fontColor: '#1F4E78' });
  s.merge('A1', 'F1');
  s.setRowHeight(0, 30);

  s.cell('A2').value = 'Fiscal year 2026 · generated in the browser';
  s.cell('A2').setStyle({ italic: true, fontColor: '#5C6670' });
  s.appendRow(['Quarter', 'Revenue', 'Cost', 'Profit', 'Margin', 'Closed on']);
  const data = [
    ['Q1', 45000, 28000, new Date(2026, 2, 31)],
    ['Q2', 62000, 31000, new Date(2026, 5, 30)],
    ['Q3', 58000, 29000, new Date(2026, 8, 30)],
    ['Q4', 85000, 39000, new Date(2026, 11, 31)],
  ];
  data.forEach(([q, rev, cost, closed], i) => {
    const r = i + 4;
    s.appendRow([q, rev, cost, `=B${r}-C${r}`, `=D${r}/B${r}`, closed]);
  });
  s.appendRow(['Total', '=SUM(B4:B7)', '=SUM(C4:C7)', '=SUM(D4:D7)', '=D8/B8', null]);

  ['A', 'B', 'C', 'D', 'E', 'F'].forEach((col) => {
    s.cell(`${col}3`).setStyle(header);
    for (let r = 4; r <= 8; r++) {
      const isTotal = r === 8;
      const numberFormat = col === 'E' ? 10 : col === 'F' ? 14 : col === 'A' ? 0 : '$#,##0';
      s.cell(`${col}${r}`).setStyle({
        bold: isTotal,
        numberFormat,
        backgroundColor: isTotal ? '#DDEBF7' : r % 2 ? '#F2F2F2' : '#FFFFFF',
        border: { style: 'thin', color: '#BFBFBF' },
      });
    }
  });
  s.cell('A8').comment = 'Formulas are recalculated when the file is opened in Excel.';
  [14, 14, 14, 14, 10, 14].forEach((w, i) => s.setColumnWidth(i, w));
  s.frozenRows = 3;
  s.tabColor = '#1F4E78';

  s.addChart({
    type: 'column',
    title: 'Revenue vs Cost',
    series: [
      { name: 'Revenue', categoriesRange: 'Dashboard!$A$4:$A$7', valuesRange: 'Dashboard!$B$4:$B$7', colorHex: '2E75B6' },
      { name: 'Cost', categoriesRange: 'Dashboard!$A$4:$A$7', valuesRange: 'Dashboard!$C$4:$C$7', colorHex: 'ED7D31' },
    ],
    anchor: { fromCol: 7, fromRow: 2, toCol: 15, toRow: 18 },
  });
  s.addChart({
    type: 'pie',
    title: 'Revenue share',
    series: [{ name: 'Revenue', categoriesRange: 'Dashboard!$A$4:$A$7', valuesRange: 'Dashboard!$B$4:$B$7' }],
    anchor: { fromCol: 7, fromRow: 19, toCol: 15, toRow: 35 },
  });
  await step('Creating image…');
  s.addImage(await pngBytes('#1F4E78', 48), 'png', 0, 10, 48, 48);
  s.cell('B11').setHyperlink('https://pub.dev/packages/excel_community', 'Open the Dart package', 'excel_community on pub.dev');

  const t = wb.createSheet('Team');
  t.appendRow(['Name', 'Region', 'Deals', 'Joined']);
  [
    ['Ana', 'North', 14, new Date(2021, 3, 12)],
    ['Bruno', 'South', 9, new Date(2023, 7, 1)],
    ['Carla', 'East', 21, new Date(2019, 0, 7)],
    ['Diego', 'West', 11, new Date(2024, 10, 18)],
  ].forEach((r) => t.appendRow(r));
  for (let r = 2; r <= 5; r++) t.cell(`D${r}`).setStyle({ numberFormat: 14 });
  t.addTable('A1:D5', 'TeamTable', ['Name', 'Region', 'Deals', 'Joined']);
  [12, 10, 8, 12].forEach((w, i) => t.setColumnWidth(i, w));
  t.protect('demo', { autoFilter: true, sort: true });

  // Highlights, input rules and printing for the dashboard.
  s.addConditionalFormatting('E4:E7', [
    { type: 'cellIs', operator: 'greaterThanOrEqual', value: 0.5, style: { backgroundColor: '#C6EFCE', fontColor: '#006100' } },
    { type: 'cellIs', operator: 'lessThan', value: 0.4, style: { backgroundColor: '#FFC7CE', fontColor: '#9C0006' }, priority: 2 },
  ]);
  s.addDataValidation('C4:C7', {
    type: 'decimal',
    operator: 'greaterThanOrEqual',
    value: 0,
    error: { title: 'Invalid cost', message: 'Costs cannot be negative.' },
  });
  s.addChart({
    type: 'line',
    title: 'Profit trend',
    dataLabels: { value: true },
    series: [{ name: 'Profit', categoriesRange: 'Dashboard!$A$4:$A$7', valuesRange: 'Dashboard!$D$4:$D$7', colorHex: '70AD47' }],
    anchor: { column: 0, row: 13, width: 6, height: 14 },
  });
  s.setPageSetup({ orientation: 'landscape', paperSize: 'a4', fitToWidth: 1, fitToHeight: 0 });
  s.setHeaderFooter({ header: '&CQuarterly Sales Dashboard', footer: '&RPage &P of &N' });

  await step('Building pivot table…');
  wb.sheet('Deals by region').addPivotTable({
    name: 'DealsByRegion',
    sourceSheet: 'Team',
    sourceRange: 'A1:C5',
    targetCell: 'A3',
    rows: ['Region'],
    values: [{ field: 'Deals', function: 'sum', customName: 'Total deals' }, { field: 'Deals', function: 'average', customName: 'Average deals' }],
  });

  wb.defaultSheet = 'Dashboard';
  return wb;
}

// ==================== VIEWER ====================

let viewed = null;
let viewedName = 'workbook.xlsx';

function showWorkbook(wb, name) {
  viewed = wb;
  viewedName = name;
  $('#viewer').hidden = false;
  $('#viewer-name').textContent = name;
  const tabs = $('#tabs');
  tabs.replaceChildren();
  wb.sheets.forEach((sheetName, i) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = sheetName;
    b.setAttribute('role', 'tab');
    b.onclick = () => showSheet(sheetName);
    tabs.append(b);
    if (i === 0) showSheet(sheetName);
  });
}

function columnName(i) {
  let s = '';
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

function showSheet(name) {
  [...$('#tabs').children].forEach((b) => b.setAttribute('aria-selected', String(b.textContent === name)));
  const sheet = viewed.sheet(name);
  const MAX_ROWS = 200;
  const MAX_COLS = 26;
  const rows = sheet.rows.slice(0, MAX_ROWS);
  const cols = Math.min(MAX_COLS, Math.max(1, ...rows.map((r) => r.length)));

  const table = document.createElement('table');
  const head = table.createTHead().insertRow();
  head.append(document.createElement('th'));
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.textContent = columnName(c);
    head.append(th);
  }
  const body = table.createTBody();
  rows.forEach((row, r) => {
    const tr = body.insertRow();
    const th = document.createElement('th');
    th.textContent = r + 1;
    tr.append(th);
    for (let c = 0; c < cols; c++) {
      const td = tr.insertCell();
      const cell = row[c];
      if (!cell || cell.value === null) continue;
      // Formulas have no cached result until Excel recalculates the file.
      td.textContent = cell.displayText || (cell.formula ? cell.formula : '');
      if (!cell.displayText && cell.formula) td.className = 'formula';
      td.title = `${cell.cellId} · ${cell.type}` + (cell.formula ? ` · ${cell.formula}` : '');
      if (cell.type === 'int' || cell.type === 'double') td.className = 'num';
      const st = cell.style;
      if (st && st.bold) td.style.fontWeight = '600';
    }
  });

  const info = [
    `${sheet.maxRows} rows × ${sheet.maxColumns} columns`,
    sheet.spannedItems.length ? `merged: ${sheet.spannedItems.join(', ')}` : null,
    sheet.frozenRows ? `frozen rows: ${sheet.frozenRows}` : null,
    sheet.hasAutoFilter ? 'auto filter' : null,
    sheet.rows.length > MAX_ROWS ? `showing first ${MAX_ROWS} rows` : null,
  ].filter(Boolean);
  $('#sheet-info').textContent = info.join(' · ');
  $('#grid').replaceChildren(table);
}

// ==================== PLAYGROUND ====================

const sample = `// \`Excel\` and \`log\` are in scope. Async code is allowed.
const wb = Excel.create();
const sheet = wb.sheet('Playground');

sheet.appendRow(['Item', 'Price']);
sheet.appendRow(['Coffee', 3.5]);
sheet.appendRow(['Cake', 4.25]);
sheet.cell('B4').formula = '=SUM(B2:B3)';
sheet.cell('A1').setStyle({ bold: true, backgroundColor: '#FFF2CC' });

log('sheets:', wb.sheets);
log('rows:', sheet.rows.map(r => r.map(c => c && c.value)));

// Uncomment to download the result:
// wb.save('playground.xlsx');
return wb;`;

function format(v) {
  if (typeof v === 'string') return v;
  try {
    return JSON.stringify(v);
  } catch (e) {
    return String(v);
  }
}

async function runPlayground() {
  const out = $('#play-output');
  out.textContent = '';
  out.classList.remove('error');
  const log = (...args) => (out.textContent += args.map(format).join(' ') + '\n');
  try {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    const result = await new AsyncFunction('Excel', 'log', $('#code').value)(Excel, log);
    if (result && typeof result.encode === 'function') {
      showWorkbook(result, 'playground.xlsx');
      log('→ returned workbook shown in the viewer below');
    }
  } catch (e) {
    out.classList.add('error');
    log(e && e.message ? e.message : String(e));
  }
}

// ==================== WIRING ====================

$('#run-tests').onclick = runTests;
$('#download-demo').onclick = async () => {
  const status = $('#demo-status');
  try {
    const wb = await buildDemo((text) => showStep(status, text));
    await downloadWorkbook(wb, 'excel-community-demo.xlsx', status);
  } catch (e) {
    showError(status, e);
  }
};
$('#view-demo').onclick = async () => {
  try {
    showWorkbook(await buildDemo(), 'excel-community-demo.xlsx');
  } catch (e) {
    showError($('#demo-status'), e);
  }
};
$('#file').onchange = async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    showWorkbook(Excel.read(new Uint8Array(await file.arrayBuffer())), file.name);
  } catch (err) {
    alert(`Could not read ${file.name}: ${err.message || err}`);
  }
};
$('#resave').onclick = async () => {
  if (!viewed) return;
  try {
    await downloadWorkbook(viewed, viewedName.replace(/\.xlsx$/i, '') + '-resaved.xlsx', $('#viewer-status'));
  } catch (e) {
    showError($('#viewer-status'), e);
  }
};
$('#code').value = sample;
$('#run-code').onclick = runPlayground;

runTests();
