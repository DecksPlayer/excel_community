![Logo](https://raw.githubusercontent.com/Decksplayer/excel_community/main/assets/logo.png)

# Excel Community (NPM Module)

> **Read and write Excel `.xlsx` files with styles, charts, tables and pivot tables, in Node.js and browsers.**  
> This package is the [excel_community](https://github.com/decksplayer/excel_community) Dart engine compiled to JavaScript, with a JS API and TypeScript types on top. No native binaries and no dependencies.

[![npm version](https://img.shields.io/npm/v/excel-community.svg)](https://www.npmjs.com/package/excel-community)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-blue.svg)](https://www.typescriptlang.org/)
[![Node.js](https://img.shields.io/badge/Node.js-%3E%3D18.0.0-green.svg)](https://nodejs.org/)
[![Bundle Size](https://img.shields.io/badge/Bundle%20Size-~216%20KB%20(gz)-brightgreen.svg)](https://www.npmjs.com/package/excel-community)

---

## 📑 Table of Contents

1. [Why Excel Community?](#-why-excel-community)
2. [Installation](#-installation)
3. [Importing the Library](#-importing-the-library)
4. [Architecture & How It Works](#-architecture--how-it-works)
5. [Quick Start (30 Seconds)](#-quick-start-30-seconds)
6. [Workbooks (`Excel` & `Workbook`)](#-workbooks-excel--workbook)
   - [Creating a Workbook](#creating-a-workbook)
   - [Reading from Disk (Node.js)](#reading-from-disk-nodejs)
   - [Reading from Buffer / Uint8Array](#reading-from-buffer--uint8array-serverless-aws-lambda)
   - [Reading in the Browser](#reading-in-the-browser-file-picker)
   - [Managing Sheets](#managing-sheets-create-rename-copy-delete)
   - [Saving & Exporting](#saving--exporting)
7. [Worksheets (`Sheet`)](#-worksheets-sheet)
   - [Cell Coordinates & Addressing](#cell-coordinates--addressing)
   - [Reading All Rows & Cells (`sheet.rows`)](#reading-all-rows--cells-sheetrows)
   - [Batch Row Appending (`appendRow`, `appendRows`)](#batch-row-appending-appendrow-appendrows)
   - [Inserting, Removing & Clearing Rows and Columns](#inserting-removing--clearing-rows-and-columns)
   - [Ranges and Find & Replace](#ranges-and-find--replace)
   - [Dimensions (Width, Height, Auto-Fit)](#dimensions-width-height-auto-fit)
   - [Freeze Panes (Lock Rows & Columns)](#freeze-panes-lock-rows--columns)
   - [Hidden Rows & Hidden Columns](#hidden-rows--hidden-columns)
   - [Merging and Unmerging Cells](#merging-and-unmerging-cells)
   - [Tab Colors & Right-To-Left (RTL)](#tab-colors--right-to-left-rtl)
   - [AutoFilters](#autofilters)
   - [Sheet Protection & Password Hashing](#sheet-protection--password-hashing)
   - [Row & Column Grouping (Outlining)](#row--column-grouping-outlining)
   - [Embedding Images](#embedding-images)
   - [Adding Charts](#adding-charts-11-types)
   - [Structured Tables ("Format as Table")](#structured-tables-format-as-table)
   - [Conditional Formatting](#conditional-formatting)
   - [Data Validation & Dropdowns](#data-validation--dropdowns)
   - [Page Setup & Printing](#page-setup--printing)
   - [Pivot Tables](#pivot-tables)
   - [Export & Import (Objects, JSON, CSV)](#export--import-objects-json-csv)
8. [Cells (`Cell`)](#-cells-cell)
   - [Reading & Writing Values](#reading--writing-values)
   - [Data Types Supported](#data-types-supported)
   - [Formatted Display Text (`cell.displayText`)](#formatted-display-text-celldisplaytext)
   - [Formulas & Calculation](#formulas--calculation)
   - [Comments & Notes](#comments--notes)
   - [Hyperlinks (Web, Email, Internal Cells)](#hyperlinks-web-email-internal-cells)
   - [Complete Styling Engine](#complete-styling-engine-cellsetstyle)
9. [Reference Tables](#-reference-tables)
   - [Standard Number Format Codes (IDs 0–49)](#standard-number-format-codes-ids-049)
   - [Border Styles](#border-styles)
   - [Sheet Protection Permissions](#sheet-protection-permissions)
   - [Chart Types & Configuration](#chart-types--configuration)
10. [Real-World Examples](#-real-world-examples)
    - [Example 1: Executive Sales Dashboard with Chart](#example-1-executive-sales-dashboard-with-chart)
    - [Example 2: In-Memory Transform & Cloud Export (AWS Lambda)](#example-2-in-memory-transform--cloud-export-aws-lambda)
11. [Errors](#️-errors)
12. [Benchmark](#-benchmark)
13. [Limitations](#-limitations)
14. [License](#-license)

---

## ⚡ Why Excel Community?

Free JavaScript Excel packages usually leave out charts, pivot tables or styling. `excel-community` includes them:

| Feature | `excel-community` | `exceljs` | `xlsx` (SheetJS Community) |
| :--- | :---: | :---: | :---: |
| **Active Open Source Community** | ✅ **Yes** | ⚠️ Low maintenance | ⚠️ Paywalled styling |
| **Full Cell Styling & Colors** | ✅ **Included** | ✅ Included | ❌ Pro Only |
| **Charts Engine (Column, Bar, Line, Pie, etc.)** | ✅ **Included** | ❌ No native charts | ❌ Pro Only |
| **Structured Tables & TableStyles** | ✅ **Included** | ⚠️ Partial | ❌ Pro Only |
| **Freeze Panes & Split Panes** | ✅ **Included** | ✅ Included | ⚠️ Limited |
| **Sheet Protection + 16-Bit Password Hash** | ✅ **Included** | ⚠️ Basic | ❌ Pro Only |
| **Images (PNG, JPEG, WebP, GIF)** | ✅ **Included** | ✅ Included | ❌ Pro Only |
| **Legacy XLS (BIFF8) + Modern XLSX** | ✅ **Included** | ❌ XLSX only | ✅ Included |
| **Bundle Size (Minified + Gzipped)** | **~216 KB** | ~252 KB | ⚠️ ~400 KB |
| **Dependencies** | ✅ **0** | 9 | ⚠️ Multiple |

---

## 📦 Installation

```bash
# npm
npm install excel-community

# yarn
yarn add excel-community

# pnpm
pnpm add excel-community
```

---

## 🔌 Importing the Library

`excel-community` ships with dual builds supporting **CommonJS**, **ES Modules**, and full **TypeScript definitions**:

```javascript
// CommonJS (Node.js)
const { Excel } = require('excel-community');

// ES Modules / Modern Bundlers (Vite, Rollup, Webpack, Next.js)
import { Excel } from 'excel-community';

// TypeScript
import { Excel, Workbook, Sheet, Cell, CellStyleOptions } from 'excel-community';
```

Without a bundler, load the standalone browser build with a `<script>` tag. It exposes `ExcelCommunity.Excel`:

```html
<script src="https://cdn.jsdelivr.net/npm/excel-community/dist/excel_community.browser.js"></script>
<script>
  const { Excel } = ExcelCommunity;
  const wb = Excel.create();
  wb.sheet('Sheet1').cell('A1').value = 'Hello';
  wb.save('hello.xlsx'); // triggers a download
</script>
```

### Example page

[`example/`](example/) contains a browser page that runs automated tests for every feature, downloads a demo workbook, previews any `.xlsx` file and has a code playground. From this folder:

```bash
npm run build     # compile the Dart core (requires the Dart SDK)
npm run example   # then open http://localhost:8080/example/
```

---

## 🧠 Architecture & How It Works

1. **Native OpenXML Engine**: `excel-community` is built on top of Dart's high-performance OpenXML serializer. It reads, parses, and writes OpenXML `.xlsx` files adhering strictly to the ISO/IEC 29500 spreadsheet specification.
2. **Compiled with dart2js**: the workbook lives in the compiled Dart code, and `Workbook`, `Sheet` and `Cell` are JavaScript objects that call into it (no JSON serialization or workers). A `Cell` is a light handle on a position of its sheet; for bulk data, `appendRows` is still the fastest way in.
3. **Environment Polyfills Built-In**: It automatically adapts to its execution environment:
   - In **Node.js**: `wb.save('file.xlsx')` calls `fs.writeFileSync`.
   - In **Browsers**: `wb.save('file.xlsx')` creates a `Blob` and triggers a download link automatically.

---

## 🚀 Quick Start (30 Seconds)

Create a styled spreadsheet and save it in 6 lines of code:

```javascript
const { Excel } = require('excel-community');

// 1. Create a workbook
const wb = Excel.create();
const sheet = wb.sheet('Summary');

// 2. Add data and formulas
sheet.cell('A1').value = 'Total Revenue';
sheet.cell('B1').value = 125000;
sheet.cell('B2').formula = '=SUM(B1*1.21)';

// 3. Apply professional styles
sheet.cell('A1').setStyle({
  bold: true,
  fontColor: '#FFFFFF',
  backgroundColor: '#1F4E78',
  horizontalAlign: 'center',
});

// 4. Save to disk
wb.save('Revenue.xlsx');
```

---

## 📘 Workbooks (`Excel` & `Workbook`)

The `Workbook` represents the complete spreadsheet file, containing one or more worksheets.

### Creating a Workbook

```javascript
const wb = Excel.create();
```
> [!NOTE]
> Creating a new workbook automatically initializes a default worksheet named `'Sheet1'`.

---

### Reading from Disk (Node.js)

```javascript
const wb = Excel.fromFile('./path/to/spreadsheet.xlsx');
const sheet = wb.sheet('Sheet1');
```

---

### Reading from Buffer / Uint8Array (Serverless, AWS Lambda)

```javascript
const fs = require('fs');

// From Node Buffer
const buffer = fs.readFileSync('./data.xlsx');
const wb = Excel.read(buffer);

// From Uint8Array or ArrayBuffer
// const wb = Excel.read(uint8Array);

// From Base64 string
// const wb = Excel.read('UEsDBBQAAAAIA...');
```

---

### Reading in the Browser (File Picker)

```javascript
import { Excel } from 'excel-community';

document.getElementById('upload').addEventListener('change', async (event) => {
  const file = event.target.files[0];
  const arrayBuffer = await file.arrayBuffer();
  
  // Load workbook from Uint8Array
  const wb = Excel.read(new Uint8Array(arrayBuffer));
  
  // Read first sheet
  const firstSheetName = wb.sheets[0];
  const sheet = wb.sheet(firstSheetName);
  
  console.log('Cell A1:', sheet.cell('A1').value);
});
```

---

### Managing Sheets (Create, Rename, Copy, Delete)

```javascript
const wb = Excel.create();

// List all sheet names (Array of strings)
console.log(wb.sheets); // ['Sheet1']

// Rename a sheet
wb.renameSheet('Sheet1', 'Executive Summary');

// Access an existing sheet or create it if it doesn't exist
// (on a new workbook whose 'Sheet1' is still empty, 'Sheet1' is renamed instead)
const salesSheet = wb.sheet('Sales');

// Explicitly add a new sheet (never renames 'Sheet1')
const inventorySheet = wb.createSheet('Inventory');

// Link two sheets: changes to either one apply to both
wb.linkSheet('Sales_View', 'Sales');
wb.unlinkSheet('Sales_View'); // give it its own copy again

// Copy / Duplicate a sheet with all its data, styles, merges and formulas
wb.copySheet('Sales', 'Sales_Q1_Backup');

// Delete a sheet (a workbook must always retain at least 1 sheet)
wb.deleteSheet('Sales_Q1_Backup');

// Set the default active sheet shown when opening in Excel
wb.defaultSheet = 'Executive Summary';
console.log(wb.defaultSheet); // 'Executive Summary'
```

---

### Saving & Exporting

```javascript
// 1. Save to local file path (Node.js) or Trigger Browser Download
wb.save('FinancialReport.xlsx');

// 2. Get Node.js Buffer (Ideal for Express.js responses, S3 upload, etc.)
const buffer = wb.toBuffer();
// res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
// res.send(buffer);

// 3. Get raw Uint8Array
const uint8Array = wb.encode();

// 4. Get Base64 encoded string
const base64 = wb.encodeBase64();
```

---

## 📄 Worksheets (`Sheet`)

A `Sheet` represents an individual tab/worksheet.

### Cell Coordinates & Addressing

Cells are indexed using standard alphanumeric A1-notation (e.g. `'A1'`, `'B12'`, `'AA50'`):
- Columns: `'A'` (0), `'B'` (1), ... `'Z'` (25), `'AA'` (26), etc.
- Rows: `'1'` (0), `'2'` (1), etc.

```javascript
const cell = sheet.cell('B5');
console.log(`Column: ${cell.col}, Row: ${cell.row}`); // Column: 1, Row: 4
```

---

### Reading All Rows & Cells (`sheet.rows`)

`sheet.rows` returns a **2D Array** of `Cell` objects representing the entire sheet grid:

```javascript
const rows = sheet.rows;

for (let r = 0; r < rows.length; r++) {
  const row = rows[r];
  for (let c = 0; c < row.length; c++) {
    const cell = row[c];
    if (cell && cell.value !== null) {
      console.log(`[${cell.cellId}] Value: ${cell.value} | Display: ${cell.displayText}`);
    }
  }
}
```

---

### Batch Row Appending (`appendRow`, `appendRows`)

Insert complete rows of data in a single call. Values can be strings, numbers, booleans, dates, or formula strings:

```javascript
// Append header row
sheet.appendRow(['ID', 'Product', 'Quantity', 'Unit Price', 'Total']);

// Append data rows
sheet.appendRow([101, 'Mechanical Keyboard', 2, 85.50, '=C2*D2']);
sheet.appendRow([102, 'Gaming Mouse', 5, 45.00, '=C3*D3']);
sheet.appendRow([103, 'USB-C Cable', 10, 8.99, '=C4*D4']);

// Many rows in one call: the fastest way to load data (20,000 rows × 10 columns take about 0.4 s in Node)
sheet.appendRows(records.map((r) => [r.id, r.product, r.quantity, r.price]));

// Query current sheet boundary dimensions
console.log(`Max Rows: ${sheet.maxRows}, Max Columns: ${sheet.maxColumns}`);
```

---

### Inserting, Removing & Clearing Rows and Columns

Indexes are 0-based. Merged cells, tables, validations, links, groups, row heights and column widths move with their cells:

```javascript
sheet.insertRow(1);       // new empty row 2; rows below move down
sheet.removeRow(5);       // row 6 is deleted; rows below move up
sheet.insertColumn(0);    // new empty column A
sheet.removeColumn(3);    // column D is deleted
sheet.clearRow(2);        // empties row 3 and keeps it in place

// Write values into an existing row, starting at a column
sheet.insertRowIterables(['Q1', 100, 200], 7, { startingColumn: 1 }); // B8:D8
```

---

### Ranges and Find & Replace

```javascript
// Values of a range as a 2D array
sheet.rangeValues('A1:C3'); // [['Region', 'Units', 'Price'], ['North', 10, 2.5], ...]

// Replace text in every text cell; returns the number of replacements
sheet.findAndReplace('Flutter', 'Dart');
sheet.findAndReplace(/draft/i, 'Final', { startingRow: 1, endingRow: 99, first: 10 });
```

---

### Dimensions (Width, Height, Auto-Fit)

```javascript
// Set Column A (index 0) width to 30 characters
sheet.setColumnWidth(0, 30);
console.log(sheet.getColumnWidth(0)); // 30

// Set Row 1 (index 0) height to 45 points
sheet.setRowHeight(0, 45);
console.log(sheet.getRowHeight(0)); // 45

// Set default dimensions for any un-configured cell
sheet.setDefaultColumnWidth(14);
sheet.setDefaultRowHeight(20);

// Enable Column Auto-Fit (computes width based on content)
sheet.setColumnAutoFit(1); // Column B
```

---

### Freeze Panes (Lock Rows & Columns)

Keep header rows and identifier columns pinned while scrolling through large datasets:

```javascript
// Lock the top row (Row 1) and the first column (Column A)
sheet.frozenRows = 1;
sheet.frozenColumns = 1;

// Lock only the first 2 header rows (no column freeze)
// sheet.frozenRows = 2;
// sheet.frozenColumns = null;

// Remove freeze panes completely
// sheet.frozenRows = null;
// sheet.frozenColumns = null;
```

---

### Hidden Rows & Hidden Columns

Hide confidential calculations, internal IDs, or reference tables:

```javascript
// Hide Column B (index 1)
sheet.setColumnHidden(1, true);
console.log(sheet.isColumnHidden(1)); // true

// Unhide Column B
sheet.setColumnHidden(1, false);

// Hide Row 5 (index 4)
sheet.setRowHidden(4, true);
console.log(sheet.isRowHidden(4)); // true
```

---

### Merging and Unmerging Cells

```javascript
// Merge banner across A1:E1
sheet.cell('A1').value = 'Annual Financial Statement';
sheet.merge('A1', 'E1');

// Or merge and set the content in one call
sheet.merge('A3', 'E3', 'Quarterly Breakdown');

// Inspect merged areas in this sheet
console.log(sheet.spannedItems); // ['A1:E1', 'A3:E3']

// Unmerge by range or by any cell inside it
sheet.unmerge('A1:E1');
sheet.unmerge('B3');
```

---

### Tab Colors & Right-To-Left (RTL)

```javascript
// Assign color to sheet tab (shown at bottom of Excel / Sheets)
sheet.tabColor = '#E63946'; // Red
// sheet.tabColor = null;   // Remove color

// Or an Office theme color (0-11) with an optional tint (-1 to 1)
sheet.setTabColorTheme(4, 0.4);

// Enable Right-To-Left (RTL) reading order (Arabic, Hebrew)
sheet.rightToLeft = true;
console.log(sheet.rightToLeft); // true
```

---

### AutoFilters

Enable Excel filter dropdowns across a range:

```javascript
// Apply AutoFilter on range A1:E50
sheet.setAutoFilter('A1:E50');
console.log(sheet.hasAutoFilter); // true

// Active criteria; `column` is the 0-based offset inside the range
sheet.addFilterColumn({ column: 1, values: ['Completed', 'In Progress'] });
sheet.addFilterColumn({ column: 3, custom: [{ operator: 'greaterThan', value: 1000 }] });
console.log(sheet.autoFilter); // { ref: 'A1:E50', columns: [...] }

// Remove AutoFilter
sheet.clearAutoFilter();
console.log(sheet.hasAutoFilter); // false
```

---

### Sheet Protection & Password Hashing

Lock worksheets with password protection and configure granular user permissions:

```javascript
// Protect sheet with password. Each option set to true ALLOWS that action
// while the sheet is protected; anything left out is blocked.
sheet.protect('SecurePassword123', {
  formatColumns: true,      // Allow resizing columns
  autoFilter: true,         // Allow using AutoFilter
  sort: true,               // Allow sorting
  selectUnlockedCells: true,
  selectLockedCells: false, // Only unlocked cells can be selected
});
console.log(sheet.isProtected); // true

// Cells are locked by default; unlock the ones users may edit,
// or hide a formula from the formula bar
sheet.cell('B2').setStyle({ locked: false });
sheet.cell('C2').setStyle({ hidden: true });

// Remove protection
sheet.unprotect();
```

---

### Row & Column Grouping (Outlining)

Create nested, collapsible hierarchical outlines (up to 7 levels):

```javascript
// Group detail rows 2 through 10, with a nested group that starts collapsed
sheet.groupRows(1, 9);
sheet.groupRows(2, 4, { collapsed: true });

// Group columns B through D
sheet.groupColumns(1, 3);

// Expand / collapse later
sheet.expandRowGroup(2, 4);
sheet.collapseColumnGroup(1, 3);

// Inspect
console.log(sheet.rowGroups);             // [{ start: 1, end: 9, level: 1, collapsed: false }, ...]
console.log(sheet.getRowOutlineLevel(3)); // 2

// Summary rows above the details (Excel's default is below)
sheet.outlineSettings = { summaryBelow: false };

// Ungroup one level, or remove every group
sheet.ungroupRows(2, 4);
sheet.clearGrouping();
```

---

### Embedding Images

Embed PNG, JPEG, GIF, or WebP images into worksheets with pixel-precise anchor positioning:

```javascript
const fs = require('fs');

const imageBuffer = fs.readFileSync('./company_logo.png');

sheet.addImage(
  imageBuffer, // Buffer or Uint8Array
  'png',       // Format: 'png' | 'jpeg' | 'gif' | 'webp'
  0,           // Column anchor (0 = Column A)
  0,           // Row anchor (0 = Row 1)
  220,         // Width in pixels
  80,          // Height in pixels
  { colOffset: 5, rowOffset: 5 } // Optional offset inside the cell, in pixels
);
```

---

### Adding Charts (11 Types)

`excel-community` includes an advanced native chart builder producing OpenXML `<c:chartSpace>` definitions:

```javascript
// 1. Prepare data
sheet.appendRow(['Quarter', 'Revenue', 'Operating Cost']);
sheet.appendRow(['Q1', 45000, 28000]);
sheet.appendRow(['Q2', 62000, 31000]);
sheet.appendRow(['Q3', 58000, 29000]);
sheet.appendRow(['Q4', 85000, 39000]);

// 2. Add a Column Chart (a config object or its JSON string)
sheet.addChart({
  type: 'column', // see Chart Types & Configuration below
  title: '2026 Financial Overview',
  showLegend: true,
  grouping: 'clustered', // 'clustered' | 'stacked' | 'percentStacked'
  series: [
    {
      name: 'Revenue',
      categoriesRange: 'Sheet1!$A$2:$A$5',
      valuesRange: 'Sheet1!$B$2:$B$5',
      colorHex: '2E75B6', // Custom hex color
    },
    {
      name: 'Operating Cost',
      categoriesRange: 'Sheet1!$A$2:$A$5',
      valuesRange: 'Sheet1!$C$2:$C$5',
      colorHex: 'ED7D31',
    },
  ],
  anchor: {
    fromCol: 4,  // Starts at column E
    fromRow: 1,  // Starts at row 2
    toCol: 13,   // Extends to column N
    toRow: 16,   // Extends to row 17
  },
});

// 3. Data labels, series styles and the short anchor form
sheet.addChart({
  type: 'line',
  title: 'Revenue trend',
  smooth: true,
  dataLabels: { value: true, labelPosition: 't' },
  series: [{
    name: 'Revenue',
    categoriesRange: 'Sheet1!$A$2:$A$5',
    valuesRange: 'Sheet1!$B$2:$B$5',
    style: { fillColor: '#2E86AB', fillType: 'transparent', fillAlpha: 40, borderColor: '#1A5276', borderWidth: 28575 },
  }],
  anchor: { column: 4, row: 18, width: 9, height: 15 }, // in cells
});

// Pie / doughnut: percentages on each slice
sheet.addChart({
  type: 'doughnut',
  dataLabels: { value: true, percentage: true, separator: '\n' },
  series: [{ name: 'Share', categoriesRange: 'Sheet1!$A$2:$A$5', valuesRange: 'Sheet1!$B$2:$B$5' }],
});

// Bubble: a third range with the bubble sizes
sheet.addChart({
  type: 'bubble',
  series: [{ name: 'Deals', categoriesRange: 'Sheet1!$B$2:$B$5', valuesRange: 'Sheet1!$C$2:$C$5', bubbleSizeRange: 'Sheet1!$D$2:$D$5' }],
});

// Stock: open, high, low and close series in that order
sheet.addChart({
  type: 'stock',
  series: ['B', 'C', 'D', 'E'].map((c) => ({ categoriesRange: 'Prices!$A$2:$A$30', valuesRange: `Prices!$${c}$2:$${c}$30` })),
});
```

> [!NOTE]
> Pie, doughnut and of-pie charts without custom colors take their slice colors from a fixed palette, in order, so saving the same workbook twice gives the same file.

---

### Structured Tables ("Format as Table")

Format ranges with native Excel TableML, automatic banded stripes, filter buttons, and column names:

```javascript
sheet.addTable(
  'A1:C5',        // Table range
  'Financials',   // Unique table identifier
  ['Quarter', 'Revenue', 'Operating Cost'] // Column titles (array or JSON string)
);

// Style, options and a totals row (labels and SUBTOTAL formulas are filled in)
sheet.addTable('A1:C10', 'Q2_Sales', [
  { name: 'Region', totalsLabel: 'Total' },
  { name: 'Units', totalsFunction: 'sum' },   // sum | average | count | countNumbers | min | max | stdDev | variance
  { name: 'Price', totalsFunction: 'average' },
], { style: 'medium9', showTotalsRow: true, showColumnStripes: true });

// Read, grow, change and remove tables
sheet.tableRowsAsMaps('Q2_Sales');                  // [{ Region: 'North', Units: 10, Price: 2.5 }, ...]
sheet.appendTableRow('Q2_Sales', ['East', 4, 1.5]);  // the table grows by one row
sheet.updateTable('Q2_Sales', { style: 'light9', showRowStripes: false });
console.log(sheet.getTable('Q2_Sales'), sheet.tables);
sheet.removeTable('Q2_Sales');                       // cells keep their values
```

Table names must be unique in the workbook and cannot contain spaces or look like cell references (`'T2'` is rejected).

With `showTotalsRow: true`, the last row of the range is the totals row, so leave it empty: if it holds data, `addTable` throws an `ExcelArgumentError` instead of overwriting it. Turning the totals row on later with `updateTable(name, { showTotalsRow: true })` adds a row below the table, like Excel.

---

### Conditional Formatting

Highlight cells with native Excel rules. Add one rule object or an array of rules to a range:

```javascript
// Compare with a value or formula
sheet.addConditionalFormatting('C2:C100', {
  type: 'cellIs',
  operator: 'greaterThan', // equal | notEqual | greaterThan | greaterThanOrEqual | lessThan | lessThanOrEqual | between | notBetween
  value: 1000,
  style: { backgroundColor: '#C6EFCE', fontColor: '#006100', bold: true },
});

sheet.addConditionalFormatting('D2:D100', [
  { type: 'cellIs', operator: 'between', value: 0, value2: 50, style: { fontColor: '#9C0006' } },
  // A formula relative to the first cell of the range
  { type: 'expression', formula: '$E2="Late"', style: { italic: true, strikethrough: true }, priority: 2 },
]);

// Text, duplicate and unique values
sheet.addConditionalFormatting('A2:A100', { type: 'containsText', text: 'urgent', style: { bold: true } });
sheet.addConditionalFormatting('B2:B100', { type: 'duplicateValues', style: { backgroundColor: '#FFC7CE' } });
sheet.addConditionalFormatting('B2:B100', { type: 'uniqueValues', style: { underline: 'single' } });

console.log(sheet.conditionalFormattings); // [{ range: 'C2:C100', rules: [...] }, ...]
sheet.clearConditionalFormatting();
```

---

### Data Validation & Dropdowns

Restrict what users can type, with optional input and error messages:

```javascript
// Dropdown with fixed values
sheet.addDataValidation('C2:C100', {
  type: 'list',
  items: ['Open', 'In progress', 'Done'],
  prompt: { title: 'Status', message: 'Pick a status from the list' },
  error: { title: 'Invalid status', message: 'Choose one of the listed values', style: 'stop' }, // stop | warning | information
});

// Dropdown from a range, on the same or another sheet
sheet.addDataValidation('D2:D100', { type: 'listFromRange', range: 'A1:A20', sheet: 'Lookup Lists' });

// Numbers, dates, times and text length
sheet.addDataValidation('E2:E100', { type: 'wholeNumber', operator: 'between', value: 1, value2: 10 });
sheet.addDataValidation('F2:F100', { type: 'decimal', operator: 'greaterThan', value: 0 });
sheet.addDataValidation('G2:G100', { type: 'date', operator: 'greaterThanOrEqual', value: new Date(2026, 0, 1) });
sheet.addDataValidation('H2:H100', { type: 'time', operator: 'lessThan', value: '18:00' });
sheet.addDataValidation('I2:I100', { type: 'textLength', operator: 'lessThanOrEqual', value: 50 });

// Custom formula (without '='), relative to the first cell of the range
sheet.addDataValidation('J2:J100', { type: 'custom', formula: 'J2>E2' });

// Check values in JavaScript before writing them
sheet.cell('C5').validates('Done'); // true
sheet.cell('E5').validates(42);     // false
console.log(sheet.getDataValidation('C5'), sheet.dataValidations);

sheet.removeDataValidation('C50:C100'); // other cells keep the rule
sheet.clearDataValidations();
```

A cell has one validation: adding a rule to a range removes that area from earlier rules.

---

### Page Setup & Printing

```javascript
// Orientation, paper, scaling (only the given options change)
sheet.setPageSetup({
  orientation: 'landscape',      // portrait | landscape
  paperSize: 'a4',               // letter | legal | a3 | a4 | a5 | ... or Excel's numeric code
  fitToWidth: 1, fitToHeight: 0, // all columns on one page width (or use `scale: 80`)
  firstPageNumber: 3,
  blackAndWhite: true,
  copies: 2,
});

// Margins: presets or values in inches / centimetres
sheet.setPageMargins('narrow');                          // normal | wide | narrow
sheet.setPageMargins({ left: 2, right: 2, unit: 'cm' });

// Print options
sheet.setPrintOptions({ gridLines: true, headings: true, horizontalCentered: true });

// Header and footer with Excel codes: &L &C &R sections, &P page, &N pages, &D date, &A sheet name
sheet.setHeaderFooter({ header: '&CQuarterly report', footer: '&LConfidential&RPage &P of &N' });

console.log(sheet.pageSetup, sheet.pageMargins, sheet.printOptions, sheet.headerFooter);
sheet.clearPageSetup();
sheet.clearPageMargins();
sheet.clearPrintOptions();
sheet.clearHeaderFooter();
```

---

### Pivot Tables

Summarize a data range on another sheet. The pivot table is saved already computed (its cells hold the results), so it also shows in Protected View and in other spreadsheet apps; Excel refreshes it when the file is opened:

```javascript
wb.sheet('Report').addPivotTable({
  name: 'SalesByRegion',
  sourceSheet: 'Sales Data',
  sourceRange: 'A1:C100',   // headers included
  targetCell: 'A3',         // top-left corner on 'Report'
  rows: ['Region'],         // header names
  columns: ['Product'],
  values: [
    { field: 'Amount', function: 'sum', customName: 'Total Sales' },
    { field: 'Amount', function: 'average' }, // sum | count | average | max | min | product | countNums | stdDev | stdDevp | var | varp
  ],
});

// The results are in the cells right away
wb.sheet('Report').rangeValues('A3:D10');

// After changing the source data (saving does this too)
wb.sheet('Report').refreshPivotTables();
```

---

### Export & Import (Objects, JSON, CSV)

The first row (or `headerRow`) gives the keys:

```javascript
// Rows as objects with native values (numbers, booleans, Date objects)
sheet.rowsAsMaps();                        // [{ Name: 'Ana', Age: 31, Joined: Date }, ...]
sheet.rowsAsMaps({ mode: 'displayText' }); // values as Excel displays them
sheet.rowsAsMaps({ headerRow: 1 });        // header below a title row

// The whole grid, header included
sheet.rowsAsValues();

// JSON and CSV (RFC 4180 quoting; CSV uses the displayed text by default)
const json = sheet.toJson({ indent: '  ' });
const csv = sheet.toCsv({ separator: ';' });

// Every sheet at once
wb.toMaps(); // { Customers: [...], Orders: [...] }
wb.toJson();

// Import objects as rows: the header is written on an empty sheet, matched otherwise
wb.sheet('Imported').appendRowsFromMaps([
  { Name: 'Eva', Age: 40, Joined: new Date(2026, 0, 15) },
]);
```

---

## 🔲 Cells (`Cell`)

Every cell in a worksheet is represented by a `Cell` instance.

### Reading & Writing Values

```javascript
const cell = sheet.cell('B2');

// Direct assignment
cell.value = 'Quarterly Report';
cell.value = 1500;
cell.value = true;
cell.value = null; // Clears the cell
```

---

### Data Types Supported

`excel-community` natively differentiates and preserves all core Excel datatypes:

```javascript
// String / Text
sheet.cell('A1').value = 'Products';
console.log(sheet.cell('A1').type); // 'string'

// Integer
sheet.cell('A2').value = 42;
console.log(sheet.cell('A2').type); // 'int'

// Double / Float
sheet.cell('A3').value = 1299.95;
console.log(sheet.cell('A3').type); // 'double'

// Boolean
sheet.cell('A4').value = true;
console.log(sheet.cell('A4').type); // 'bool'

// Dates & DateTimes (Date objects, stored with their local time)
sheet.cell('A5').value = new Date(2026, 9, 4);  // type 'date', value '2026-10-04'
sheet.cell('A6').value = new Date();            // type 'datetime', ISO string value
console.log(sheet.cell('A5').dateValue);        // a JS Date again

// Times of day
sheet.cell('A7').setTime('14:30');              // type 'time', value '14:30:00'
sheet.cell('A8').setTime(8, 15, 0);
```

---

### Formatted Display Text (`cell.displayText`)

`cell.displayText` provides the **rendered text** as it appears on the screen in Excel according to its assigned `numberFormat`:

```javascript
sheet.cell('B2').value = 1250.5;
sheet.cell('B2').setStyle({ numberFormat: '$#,##0.00' });

console.log(sheet.cell('B2').value);       // 1250.5 (numeric)
console.log(sheet.cell('B2').displayText); // "$1,250.50" (rendered string)
```

---

### Formulas & Calculation

```javascript
// Assign formula via property
sheet.cell('C1').formula = '=SUM(A1:B1)';

// Or assign value starting with "="
sheet.cell('C2').value = '=AVERAGE(A1:B1)';

console.log(sheet.cell('C1').formula); // '=SUM(A1:B1)'
console.log(sheet.cell('C1').type);    // 'formula'

// For files saved by Excel: the result Excel cached for the formula
console.log(Excel.fromFile('report.xlsx').sheet('Sheet1').cell('C1').cachedValue);
```

Formulas are not calculated by the library; Excel calculates them when the file is opened.

---

### Comments & Notes

Attach rich review comments (displaying a red corner triangle indicator in Microsoft Excel):

```javascript
sheet.cell('B2').comment = 'Audited by Finance Dept on Oct 2026';
console.log(sheet.cell('B2').comment); // 'Audited by Finance Dept on Oct 2026'

// Remove comment
sheet.cell('B2').comment = null;
```

---

### Hyperlinks (Web, Email, Internal Cells)

```javascript
// Link to web page
sheet.cell('A1').setHyperlink(
  'https://github.com/decksplayer/excel_community',
  'Open GitHub Repository', // Tooltip hover text
  'Visit excel_community'   // Display text
);

// Or pass a target object: url, email, a cell of the workbook or a defined name
sheet.cell('A2').setHyperlink({ email: 'sales@example.com', subject: 'Q3 report' }, { text: 'Contact sales' });
sheet.cell('A3').setHyperlink({ sheet: 'Q1 Sales', cell: 'B4' }, { text: 'Go to Q1' });
sheet.cell('A4').setHyperlink({ location: 'TotalSales', tooltip: 'Named range' });
sheet.cell('A5').setHyperlink({ url: 'https://dart.dev' }, { styled: false }); // keep the cell style

// The same link over a range
sheet.setHyperlinkRange('B5:D5', { url: 'https://pub.dev' });

// Read hyperlink
const link = sheet.cell('A1').getHyperlink();
console.log(link.url);     // 'https://github.com/...'
console.log(link.tooltip); // 'Open GitHub Repository'
console.log(sheet.hyperlinks); // { A1: {...}, 'B5:D5': {...}, ... }

// Remove one or all hyperlinks
sheet.cell('A1').removeHyperlink();
sheet.clearHyperlinks();
```

---

### Complete Styling Engine (`cell.setStyle`)

`cell.setStyle()` provides full control over cell presentation. It changes only the options you pass, so you can call it several times; `cell.resetStyle()` removes all formatting:

```javascript
sheet.cell('A1').setStyle({
  // Font
  fontFamily: 'Segoe UI',
  fontSize: 14,
  bold: true,
  italic: false,
  underline: 'single', // 'single' | 'double' | true | false
  strikethrough: false,

  // Colors
  fontColor: '#FFFFFF',        // White text
  backgroundColor: '#1B365D',  // Navy Blue background

  // Alignment
  horizontalAlign: 'center',   // 'left' | 'center' | 'right'
  verticalAlign: 'center',     // 'top' | 'center' | 'bottom'
  wrapText: true,              // Wrap text onto multiple lines
  // shrinkToFit: true,        // Or shrink the text to fit the cell
  rotation: 0,                 // Text angle from -90 to 90 degrees

  // Borders
  border: { style: 'thin', color: '#B0C4DE' },
  // Or specify individual borders:
  // leftBorder:   { style: 'thin',   color: '#B0C4DE' },
  // rightBorder:  { style: 'thin',   color: '#B0C4DE' },
  // topBorder:    { style: 'thick',  color: '#002060' },
  // bottomBorder: { style: 'double', color: '#002060' },

  // Diagonal border
  // diagonalBorder: { style: 'thin', color: '#FF0000' }, diagonalUp: true, diagonalDown: false,

  // Number Format
  numberFormat: '$#,##0.00;($#,##0.00);"-"',

  // Protection (used when the sheet is protected)
  locked: true,   // read-only under protection (Excel's default)
  hidden: false,  // hide the formula from the formula bar
});

// Later calls keep the other options
sheet.cell('A1').setStyle({ italic: true });

// Inspect active styles: same form setStyle takes ('#1F4E78', 'center'), so it can be passed back
const style = sheet.cell('A1').style;
console.log(style.bold, style.fontColor, style.backgroundColor, style.horizontalAlign, style.numberFormat);
sheet.cell('B1').setStyle(style);

// Remove all formatting
sheet.cell('A1').resetStyle();
```

---

## 📊 Reference Tables

### Standard Number Format Codes (IDs 0–49)

You can pass either standard integer IDs or custom formatting pattern strings to `numberFormat`:

| ID | Format Code | Description / Example |
| :---: | :--- | :--- |
| `0` | `General` | Default standard format (`1234.56`) |
| `1` | `0` | Decimal integer (`1235`) |
| `2` | `0.00` | Fixed two decimals (`1234.56`) |
| `3` | `#,##0` | Thousands separator (`1,235`) |
| `4` | `#,##0.00` | Thousands + two decimals (`1,234.56`) |
| `9` | `0%` | Integer percentage (`15%`) |
| `10` | `0.00%` | Two decimal percentage (`15.25%`) |
| `11` | `0.00E+00` | Scientific notation (`1.23E+03`) |
| `14` | `mm-dd-yy` | Short date (`10-04-26`) |
| `15` | `d-mmm-yy` | Day-month-year (`4-Oct-26`) |
| `16` | `d-mmm` | Day-month (`4-Oct`) |
| `17` | `mmm-yy` | Month-year (`Oct-26`) |
| `18` | `h:mm AM/PM` | 12-hour time (`07:30 PM`) |
| `19` | `h:mm:ss AM/PM` | 12-hour time with seconds |
| `20` | `h:mm` | 24-hour military time (`19:30`) |
| `21` | `h:mm:ss` | 24-hour military time with seconds |
| `22` | `m/d/yy h:mm` | Combined date and time |
| `37` | `#,##0 ;(#,##0)` | Accounting integer (negatives in parens) |
| `38` | `#,##0 ;[Red](#,##0)` | Negatives in red |
| `41` | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | Currency aligned accounting |
| `42` | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)` | Currency aligned with dollar sign |
| `44` | `_($* #,##0.00_);...` | Currency aligned with decimals |
| `49` | `@` | Explicit text string |

---

### Border Styles

| Style Name | Description |
| :--- | :--- |
| `'thin'` | Solid hairline thin border (standard cell grid) |
| `'medium'` | Medium-weight solid border |
| `'thick'` | Heavy-weight bold border |
| `'double'` | Classical accounting double-underline border |
| `'dashed'` | Dashed border line |
| `'dotted'` | Dotted border line |
| `'hair'` | Ultra-fine subtle boundary line |
| `'dashDot'` | Alternating dash-and-dot border line |
| `'dashDotDot'` | Alternating dash-dot-dot border line |
| `'mediumDashed'`, `'mediumDashDot'`, `'mediumDashDotDot'`, `'slantDashDot'` | Medium-weight dashed variants |
| `'none'` | Explicitly clears border |

---

### Sheet Protection Permissions

Pass any of the following booleans inside `sheet.protect('password', options)`. `true` **allows** the action while the sheet is protected; options you leave out are blocked:

| Option | Allows the user to |
| :--- | :--- |
| `selectLockedCells` | Select locked cells |
| `selectUnlockedCells` | Select unlocked cells |
| `formatCells` | Modify fonts, fills, and alignments |
| `formatColumns` | Resize or auto-fit column widths |
| `formatRows` | Resize row heights |
| `insertColumns` | Insert new columns |
| `insertRows` | Insert new rows |
| `insertHyperlinks` | Insert hyperlinks |
| `deleteColumns` | Delete columns |
| `deleteRows` | Delete rows |
| `sort` | Sort ranges |
| `autoFilter` | Use AutoFilter |
| `pivotTables` | Use pivot tables |
| `objects` | Edit charts, images and other objects |
| `scenarios` | Edit scenarios |

---

### Chart Types & Configuration

| Chart Type | String Key | Best Used For |
| :--- | :---: | :--- |
| **Column Chart** | `'column'` | Vertical bars for category comparisons (`grouping`) |
| **Bar Chart** | `'bar'` | Horizontal bars for rankings or long label lists (`grouping`) |
| **Line Chart** | `'line'` | Time series and trends (`grouping`, `showMarkers`, `smooth`) |
| **Area Chart** | `'area'` | Cumulative volume trends over time (`grouping`) |
| **Pie Chart** | `'pie'` | Proportional composition of a whole |
| **Doughnut Chart** | `'doughnut'` | Pie with a center hole |
| **Of-Pie Chart** | `'ofPie'`, `'pieOfPie'`, `'barOfPie'` | Pie with a secondary breakdown (`ofPieType`, `splitType`, `splitPosition`, `secondPieSize`) |
| **Scatter Chart** | `'scatter'` | Correlation between two numeric variables (`showLines`, `showMarkers`, `smooth`) |
| **Bubble Chart** | `'bubble'` | X, Y and bubble size (`bubbleSizeRange` per series, `bubbleScale`, `showNegativeBubbles`) |
| **Stock Chart** | `'stock'` | High-low-close or open-high-low-close (`showHighLowLines`, `showUpDownBars`) |
| **Radar Chart** | `'radar'` | Multi-variable comparative analysis (`filled`) |

Every chart also accepts `title`, `showLegend`, `anchor`, `dataLabels` (`value`, `categoryName`, `seriesName`, `percentage`, `separator`, `labelPosition`) and a `style` per series (`fillColor`, `fillType: 'solid' | 'transparent' | 'none'`, `fillAlpha`, `borderColor`, `borderAlpha`, `borderWidth`).

---

## 💼 Real-World Examples

### Example 1: Executive Sales Dashboard with Chart

```javascript
const { Excel } = require('excel-community');

const wb = Excel.create();
const sheet = wb.sheet('Q1 Dashboard');

// 1. Header Banner
sheet.cell('A1').value = 'ACME Corp — Executive Quarterly Sales';
sheet.merge('A1', 'D1');
sheet.cell('A1').setStyle({
  bold: true,
  fontSize: 16,
  fontColor: '#FFFFFF',
  backgroundColor: '#1F4E78',
  horizontalAlign: 'center',
});
sheet.setRowHeight(0, 40);

// 2. Column Headers
const headers = ['Region', 'Units Sold', 'Unit Revenue', 'Total Revenue'];
sheet.appendRow(headers);
for (let c = 0; c < 4; c++) {
  const colLetter = String.fromCharCode(65 + c);
  sheet.cell(`${colLetter}2`).setStyle({
    bold: true,
    fontColor: '#FFFFFF',
    backgroundColor: '#2E75B6',
    horizontalAlign: 'center',
  });
}

// 3. Data Rows
sheet.appendRow(['North America', 1250, 45.0, '=B3*C3']);
sheet.appendRow(['Europe', 980, 52.0, '=B4*C4']);
sheet.appendRow(['Asia-Pacific', 1650, 38.0, '=B5*C5']);
sheet.appendRow(['Latin America', 620, 34.0, '=B6*C6']);

// Format numeric columns
for (let r = 3; r <= 6; r++) {
  sheet.cell(`B${r}`).setStyle({ numberFormat: '#,##0' });
  sheet.cell(`C${r}`).setStyle({ numberFormat: '$#,##0.00' });
  sheet.cell(`D${r}`).setStyle({ numberFormat: '$#,##0.00', bold: true });
}

// 4. Auto-fit & dimensions
sheet.setColumnWidth(0, 22);
sheet.setColumnWidth(1, 15);
sheet.setColumnWidth(2, 18);
sheet.setColumnWidth(3, 20);

// 5. Freeze Header
sheet.frozenRows = 2;

// 6. Embed Column Chart
sheet.addChart(JSON.stringify({
  type: 'column',
  title: 'Quarterly Regional Performance',
  series: [
    {
      name: 'Total Revenue',
      categoriesRange: "'Q1 Dashboard'!$A$3:$A$6",
      valuesRange: "'Q1 Dashboard'!$D$3:$D$6",
      colorHex: '1F4E78',
    },
  ],
  anchor: { fromCol: 5, fromRow: 1, toCol: 13, toRow: 16 },
}));

wb.save('SalesDashboard.xlsx');
console.log('Generated SalesDashboard.xlsx successfully!');
```

---

### Example 2: In-Memory Transform & Cloud Export (AWS Lambda)

```javascript
const { Excel } = require('excel-community');

exports.handler = async (event) => {
  // 1. Read input spreadsheet from S3/API payload
  const incomingBuffer = Buffer.from(event.body, 'base64');
  const wb = Excel.read(incomingBuffer);
  
  const sheet = wb.sheet(wb.sheets[0]);
  
  // 2. Iterate rows and calculate tax
  const rows = sheet.rows;
  for (let r = 1; r < rows.length; r++) {
    const subtotal = sheet.cell(`C${r + 1}`).value;
    if (typeof subtotal === 'number') {
      sheet.cell(`D${r + 1}`).value = subtotal * 0.21; // 21% Tax
      sheet.cell(`E${r + 1}`).formula = `=C${r + 1}+D${r + 1}`;
    }
  }

  // 3. Return updated workbook directly as Base64
  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Processed.xlsx"',
    },
    body: wb.encodeBase64(),
    isBase64Encoded: true,
  };
};
```

---

## ⚠️ Errors

Errors are thrown as these classes (exported next to `Excel`), with the original error in `cause`:

| Class | When |
| :--- | :--- |
| `ExcelArgumentError` | An invalid argument or option: cell reference, range, table name, style value... |
| `ExcelStateError` | The operation is not possible in the current state. |
| `ExcelFormatError` | `Excel.read` / `Excel.fromFile` got data that is not a readable workbook. |
| `ExcelError` | Base class of the three above. |

```javascript
import { Excel, ExcelArgumentError, ExcelFormatError } from 'excel-community';

try {
  wb = Excel.read(upload);
} catch (e) {
  if (e instanceof ExcelFormatError) return res.status(400).send('Not an Excel file');
  throw e;
}
```

---

## 📈 Benchmark

`npm run bench` builds a sheet with 10 mixed columns (numbers, text, booleans) and encodes it, then reads the same file back and walks every value. Each measurement runs in a fresh Node process with the libraries alternating, and the time excludes loading the module. Node 22 on Windows (Ryzen 7 5700U laptop), median of 5 runs, against exceljs 4.4.0:

| Rows | Library | Write | Read | File size |
| ---: | :--- | ---: | ---: | ---: |
| 10,000 | excel-community | 0.96 s | 1.09 s | 679 KB |
| 10,000 | exceljs | 0.85 s | 0.65 s | 537 KB |
| 50,000 | excel-community | 3.67 s | 4.21 s | 3457 KB |
| 50,000 | exceljs | 3.55 s | 2.21 s | 2662 KB |
| 100,000 | excel-community | 6.64 s | 9.56 s | 6926 KB |
| 100,000 | exceljs | 6.69 s | 4.35 s | 5320 KB |

Writing is on par with exceljs; reading takes about twice as long, and files are about 30% larger. Results depend on the machine, so run it on yours.

---

## 🚧 Limitations

- Files are tested with the library's own reader and with openpyxl; complex workbooks from third parties (macros, external links, unusual formats) can lose parts the library does not model. Keep the original when you rewrite a user's file.
- `.xls` files can be read, but saving always writes `.xlsx`.
- Formulas are stored, not calculated: Excel computes them when the file is opened (`cachedValue` holds the result Excel saved).
- The API is synchronous and works on the whole workbook in memory; there is no streaming for very large files.
- Reading takes about twice as long as with exceljs (writing is on par); see the [benchmark](#-benchmark).
- A `Cell` is a position: after `insertRow` or `removeRow`, get the cell again with `sheet.cell(...)`.

---

## 📄 License

This library is open-source under the **[MIT License](https://opensource.org/licenses/MIT)**.  
Maintained with ❤️ by [DecksPlayer](https://github.com/decksplayer).
