![Logo](https://raw.githubusercontent.com/Decksplayer/excel_community/main/assets/logo.png)

# Excel Community (NPM Module)

> **The ultimate high-performance, full-featured Excel XLSX library for JavaScript, TypeScript, Node.js, and modern Web Browsers.**  
> Compiled directly from the industry-proven [excel_community](https://github.com/decksplayer/excel_community) Dart engine. Zero heavy native binaries. Zero bloated dependencies. 100% pure, optimized OpenXML spreadsheet manipulation.

[![npm version](https://img.shields.io/npm/v/excel-community.svg)](https://www.npmjs.com/package/excel-community)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-blue.svg)](https://www.typescriptlang.org/)
[![Node.js](https://img.shields.io/badge/Node.js-%3E%3D18.0.0-green.svg)](https://nodejs.org/)
[![Bundle Size](https://img.shields.io/badge/Bundle%20Size-~192%20KB%20(gz)-brightgreen.svg)](https://www.npmjs.com/package/excel-community)

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
   - [Batch Row Appending (`sheet.appendRow`)](#batch-row-appending-sheetappendrow)
   - [Dimensions (Width, Height, Auto-Fit)](#dimensions-width-height-auto-fit)
   - [Freeze Panes (Lock Rows & Columns)](#freeze-panes-lock-rows--columns)
   - [Hidden Rows & Hidden Columns](#hidden-rows--hidden-columns)
   - [Merging and Unmerging Cells](#merging-and-unmerging-cells)
   - [Tab Colors & Right-To-Left (RTL)](#tab-colors--right-to-left-rtl)
   - [AutoFilters](#autofilters)
   - [Sheet Protection & Password Hashing](#sheet-protection--password-hashing)
   - [Row & Column Grouping (Outlining)](#row--column-grouping-outlining)
   - [Embedding Images](#embedding-images)
   - [Adding Charts](#adding-charts-7-types)
   - [Structured Tables ("Format as Table")](#structured-tables-format-as-table)
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
11. [License](#-license)

---

## ⚡ Why Excel Community?

Traditional JavaScript Excel packages present severe trade-offs: either they are abandoned, locked behind commercial paywalls with stripped features, or weigh several megabytes with dozens of transitives.

`excel-community` solves this completely:

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
| **Bundle Size (Minified + Gzipped)** | ⚡ **~192 KB** | ❌ ~2.5 MB | ⚠️ ~400 KB |
| **Zero Heavy Dependencies** | ✅ **0 deps** | ❌ 10+ dependencies | ⚠️ Multiple |

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

---

## 🧠 Architecture & How It Works

1. **Native OpenXML Engine**: `excel-community` is built on top of Dart's high-performance OpenXML serializer. It reads, parses, and writes OpenXML `.xlsx` files adhering strictly to the ISO/IEC 29500 spreadsheet specification.
2. **Direct Memory Mapping via `@JSExport`**: The compiled JavaScript runtime uses direct JavaScript Interop objects. When you access `sheet.cell('A1').value`, JavaScript reads the memory directly with $O(1)$ lookups without slow JSON serialization or message queues.
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

// Access an existing sheet or create it if it doesn't exist
const salesSheet = wb.sheet('Sales');

// Explicitly create a new sheet
const inventorySheet = wb.createSheet('Inventory');

// Rename a sheet
wb.renameSheet('Sheet1', 'Executive Summary');

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

### Batch Row Appending (`sheet.appendRow`)

Insert complete rows of data in a single call. Values can be strings, numbers, booleans, dates, or formula strings:

```javascript
// Append header row
sheet.appendRow(['ID', 'Product', 'Quantity', 'Unit Price', 'Total']);

// Append data rows
sheet.appendRow([101, 'Mechanical Keyboard', 2, 85.50, '=C2*D2']);
sheet.appendRow([102, 'Gaming Mouse', 5, 45.00, '=C3*D3']);
sheet.appendRow([103, 'USB-C Cable', 10, 8.99, '=C4*D4']);

// Query current sheet boundary dimensions
console.log(`Max Rows: ${sheet.maxRows}, Max Columns: ${sheet.maxColumns}`);
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

// Inspect merged areas in this sheet
console.log(sheet.spannedItems); // ['A1:E1']

// Unmerge
sheet.unmerge('A1');
```

---

### Tab Colors & Right-To-Left (RTL)

```javascript
// Assign color to sheet tab (shown at bottom of Excel / Sheets)
sheet.tabColor = '#E63946'; // Red
// sheet.tabColor = null;   // Remove color

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

// Remove AutoFilter
sheet.clearAutoFilter();
console.log(sheet.hasAutoFilter); // false
```

---

### Sheet Protection & Password Hashing

Lock worksheets with password protection and configure granular user permissions:

```javascript
// Protect sheet with password and custom permissions
sheet.protect('SecurePassword123', {
  formatCells: false,      // Prevent cell formatting
  formatColumns: false,    // Prevent column resizing
  formatRows: false,       // Prevent row resizing
  insertColumns: false,    // Prevent adding columns
  insertRows: false,       // Prevent adding rows
  deleteColumns: false,    // Prevent deleting columns
  deleteRows: false,       // Prevent deleting rows
  autoFilter: true,        // Allow using AutoFilter
  sort: true,              // Allow sorting
});

// Remove protection
sheet.unprotect();
```

---

### Row & Column Grouping (Outlining)

Create nested, collapsible hierarchical outlines (up to 7 levels):

```javascript
// Group detail rows 2 through 10
sheet.groupRows(1, 9);

// Group columns B through D
sheet.groupColumns(1, 3);

// Ungroup when needed
sheet.ungroupRows(1, 9);
sheet.ungroupColumns(1, 3);
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
  80           // Height in pixels
);
```

---

### Adding Charts (7 Types)

`excel-community` includes an advanced native chart builder producing OpenXML `<c:chartSpace>` definitions:

```javascript
// 1. Prepare data
sheet.appendRow(['Quarter', 'Revenue', 'Operating Cost']);
sheet.appendRow(['Q1', 45000, 28000]);
sheet.appendRow(['Q2', 62000, 31000]);
sheet.appendRow(['Q3', 58000, 29000]);
sheet.appendRow(['Q4', 85000, 39000]);

// 2. Add a Column Chart
sheet.addChart(JSON.stringify({
  type: 'column', // 'column' | 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'radar'
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
}));
```

---

### Structured Tables ("Format as Table")

Format ranges with native Excel TableML, automatic banded stripes, filter buttons, and column names:

```javascript
sheet.addTable(
  'A1:C5',        // Table range
  'Financials',   // Unique table identifier
  JSON.stringify(['Quarter', 'Revenue', 'Operating Cost']) // Column titles
);
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

// Dates & DateTimes (ISO string or Date object)
sheet.cell('A5').value = '2026-10-04';
sheet.cell('A6').value = new Date();
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
```

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

// Read hyperlink
const link = sheet.cell('A1').getHyperlink();
console.log(link.url);     // 'https://github.com/...'
console.log(link.tooltip); // 'Open GitHub Repository'

// Remove hyperlink
sheet.cell('A1').removeHyperlink();
```

---

### Complete Styling Engine (`cell.setStyle`)

`cell.setStyle()` provides full control over cell presentation:

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
  rotation: 0,                 // Text angle from -90 to 90 degrees

  // Borders
  border: { style: 'thin', color: '#B0C4DE' },
  // Or specify individual borders:
  // leftBorder:   { style: 'thin',   color: '#B0C4DE' },
  // rightBorder:  { style: 'thin',   color: '#B0C4DE' },
  // topBorder:    { style: 'thick',  color: '#002060' },
  // bottomBorder: { style: 'double', color: '#002060' },

  // Number Format
  numberFormat: '$#,##0.00;($#,##0.00);"-"',
});

// Inspect active styles
const style = sheet.cell('A1').style;
console.log(style.bold, style.fontColor, style.backgroundColor);
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
| `'dashdot'` | Alternating dash-and-dot border line |
| `'dashdotdot'` | Alternating dash-dot-dot border line |
| `'none'` | Explicitly clears border |

---

### Sheet Protection Permissions

Pass any of the following booleans inside `sheet.protect('password', options)`:

| Option | Default | Description |
| :--- | :---: | :--- |
| `formatCells` | `true` | Allows user to modify fonts, fills, and alignments |
| `formatColumns` | `true` | Allows user to resize or auto-fit column widths |
| `formatRows` | `true` | Allows user to resize row heights |
| `insertColumns` | `true` | Allows inserting new columns |
| `insertRows` | `true` | Allows inserting new rows |
| `deleteColumns` | `true` | Allows deleting columns |
| `deleteRows` | `true` | Allows deleting rows |
| `autoFilter` | `true` | Allows applying filters |
| `sort` | `true` | Allows sorting ranges |

---

### Chart Types & Configuration

| Chart Type | String Key | Best Used For |
| :--- | :---: | :--- |
| **Column Chart** | `'column'` | Vertical bars for category comparisons |
| **Bar Chart** | `'bar'` | Horizontal bars for rankings or long label lists |
| **Line Chart** | `'line'` | Continuous time-series and trend tracking |
| **Pie Chart** | `'pie'` | Proportional composition of a whole |
| **Area Chart** | `'area'` | Cumulative volume trends over time |
| **Scatter Chart** | `'scatter'` | Correlation between two numeric variables |
| **Radar Chart** | `'radar'` | Multi-variable comparative analysis |

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

## 📄 License

This library is open-source under the **[MIT License](https://opensource.org/licenses/MIT)**.  
Maintained with ❤️ by [DecksPlayer](https://github.com/decksplayer).
