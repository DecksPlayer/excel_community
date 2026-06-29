// Basic / structural snippets: simple workbooks, multi-sheet, freeze panes,
// protection / cell locking, pivot-template round-trip and full report.
library;

// ---------------------------------------------------------------------------
// Simple workbook
// ---------------------------------------------------------------------------
const String simpleExcelSnippet = '''
import 'package:excel_community/excel_community.dart';

void main() {
  // Create an Excel workbook
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Update cell values using CellIndex
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

  // Save the workbook
  var bytes = excel.save(fileName: 'simple_no_chart.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Multi-sheet
// ---------------------------------------------------------------------------
const String multiSheetsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateMultiSheets() {
  // Create a new Excel workbook
  var excel = Excel.createExcel();
  
  // Sheet 1: Summary Sheet
  var summarySheet = excel['Summary'];
  excel.delete('Sheet1'); // Remove the default worksheet

  summarySheet.updateCell(
    CellIndex.indexByString("A1"), 
    TextCellValue("Monthly Financial Summary"),
    cellStyle: CellStyle(bold: true, fontSize: 14)
  );

  // Sheet 2: Revenues Detail
  var revenueSheet = excel['Revenues'];
  revenueSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Source"));
  revenueSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"));
  revenueSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("SaaS Subscriptions"));
  revenueSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(15400));
  revenueSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Consulting"));
  revenueSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(4200));

  // Sheet 3: Expenses Detail
  var expenseSheet = excel['Expenses'];
  expenseSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
  expenseSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Cost"));
  expenseSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Hosting & Cloud"));
  expenseSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(3100));
  expenseSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Marketing"));
  expenseSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(1200));

  // Save the multi-sheet workbook
  excel.save(fileName: 'multisheet_financials.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Freeze panes (single sheet)
// ---------------------------------------------------------------------------
const String freezePanesSnippet = r'''
import 'package:excel_community/excel_community.dart';

void generateFrozenPanesReport() {
  // 1. Create a new Excel workbook and sheet
  var excel = Excel.createExcel();
  var sheet = excel['Sales Report'];
  excel.delete('Sheet1');

  // 2. Add header row
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Product Name"), cellStyle: CellStyle(bold: true));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Sales Revenue"), cellStyle: CellStyle(bold: true));

  // 3. Populate data rows
  for (int i = 2; i <= 50; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue("Product $i"));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), IntCellValue(150 * i));
  }

  // 4. Freeze Row 1 and Column A programmatically.
  //    The values are nullable: pass `null` (or `0`) to remove the freeze.
  //    On save, a <pane xSplit ySplit topLeftCell state="frozen"/> element
  //    is written; on decode, the values are restored from the same element.
  sheet.frozenRows = 1;
  sheet.frozenColumns = 1;

  // 5. Save the output
  excel.save(fileName: 'sales_report_frozen.xlsx');

  // 6. (Optional) Verify round-trip:
  //    final reopened = Excel.decodeBytes(excel.encode()!);
  //    assert(reopened['Sales Report'].frozenRows == 1);
  //    assert(reopened['Sales Report'].frozenColumns == 1);
}
''';

// ---------------------------------------------------------------------------
// Freeze panes (multi sheet)
// ---------------------------------------------------------------------------
const String multiFreezePanesSnippet = r'''
import 'package:excel_community/excel_community.dart';

void generateMultiFreezePanesReport() {
  // 1. Create a new Excel workbook with multiple sheets
  var excel = Excel.createExcel();
  excel.delete('Sheet1');

  // -------- Sheet 1: Sales — header row + product column frozen ------------
  var sales = excel['Sales'];
  _writeHeader(sales, ['Product', 'Revenue', 'Units', 'Region']);
  for (var i = 2; i <= 25; i++) {
    sales.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue('Product $i'));
    sales.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), IntCellValue(150 * i));
    sales.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1), IntCellValue(10 * i));
    sales.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1), TextCellValue('Region $i'));
  }
  sales.frozenRows = 1;
  sales.frozenColumns = 1;

  // -------- Sheet 2: Inventory — only rows frozen (2 header rows) -----------
  var inventory = excel['Inventory'];
  _writeHeader(inventory, ['SKU', 'Name', 'Category', 'Stock', 'Price']);
  for (var i = 2; i <= 18; i++) {
    inventory.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue('SKU-$i'));
    inventory.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), TextCellValue('Item $i'));
    inventory.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1), TextCellValue('Category $i'));
    inventory.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1), IntCellValue(50 * i));
    inventory.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1), DoubleCellValue(9.99 * i));
  }
  inventory.frozenRows = 2;
  inventory.frozenColumns = null;

  // -------- Sheet 3: Customers — only the first 2 columns frozen ------------
  var customers = excel['Customers'];
  _writeHeader(customers, ['AccountId', 'RegionId', 'Name', 'Email', 'Tier']);
  for (var i = 2; i <= 20; i++) {
    customers.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue('ACC-$i'));
    customers.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), TextCellValue('REG-$i'));
    customers.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1), TextCellValue('Customer $i'));
    customers.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1), TextCellValue('c$i@example.com'));
    customers.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1), TextCellValue('Tier $i'));
  }
  customers.frozenRows = null;
  customers.frozenColumns = 2;

  // -------- Sheet 4: Logs — no freeze at all --------------------------------
  var logs = excel['Logs'];
  _writeHeader(logs, ['Timestamp', 'Level', 'Message']);
  for (var i = 2; i <= 18; i++) {
    logs.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue('2026-06-$i'));
    logs.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), TextCellValue('INFO'));
    logs.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1), TextCellValue('Event #$i'));
  }
  logs.frozenRows = null;
  logs.frozenColumns = null;

  // 2. Save the workbook
  excel.save(fileName: 'multi_freeze_panes.xlsx');

  // 3. (Optional) Verify round-trip:
  //    final reopened = Excel.decodeBytes(excel.encode()!);
  //    assert(reopened['Sales'].frozenRows == 1);
  //    assert(reopened['Sales'].frozenColumns == 1);
  //    assert(reopened['Inventory'].frozenRows == 2);
  //    assert(reopened['Customers'].frozenColumns == 2);
  //    assert(reopened['Logs'].frozenRows == null);
}

void _writeHeader(Sheet sheet, List<String> headers) {
  for (var c = 0; c < headers.length; c++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
      TextCellValue(headers[c]),
      cellStyle: CellStyle(bold: true),
    );
  }
}
''';

// ---------------------------------------------------------------------------
// Sheet protection / cell locking
// ---------------------------------------------------------------------------
const String cellLockingSnippet = '''
import 'package:excel_community/excel_community.dart';

Future<void> generateLockedCellsReport() async {
  // 1. Create a new Excel workbook and sheet
  var excel = Excel.createExcel();
  var sheet = excel['Protected Report'];
  excel.delete('Sheet1');

  // 2. Protect the sheet programmatically with a password hash
  sheet.protect('password');

  // Excel locks all cells by default when sheet protection is active.
  // Explicitly configure cell locks:
  final headerStyle = CellStyle(
    bold: true,
    locked: true,
    backgroundColorHex: ExcelColor.fromHexString('#D9D9D9'),
  );
  
  // Set locked: false to allow editing on these cells
  final editableStyle = CellStyle(
    locked: false, 
    backgroundColorHex: ExcelColor.fromHexString('#E2EFDA'),
  );
  
  final readOnlyStyle = CellStyle(locked: true);

  // 3. Write data to the cells
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Read-Only Header"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Editable Values"), cellStyle: headerStyle);

  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North Sales"), cellStyle: readOnlyStyle);
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(8500), cellStyle: editableStyle);

  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("South Sales"), cellStyle: readOnlyStyle);
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(6400), cellStyle: editableStyle);

  // 4. Save the output. Sheet protection and cell lock styles are serialized successfully.
  excel.save(fileName: 'protected_sales_report.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Pivot-template round-trip
// ---------------------------------------------------------------------------
const String pivotTemplateSnippet = '''
import 'package:flutter/services.dart' show rootBundle;
import 'package:excel_community/excel_community.dart';

Future<void> generatePivotFromTemplate() async {
  // 1. Load an existing Excel template containing pre-configured Pivot Tables
  final ByteData data = await rootBundle.load('assets/financial_template.xlsx');
  final List<int> bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  
  // 2. Decode the workbook bytes using excel_community
  var excel = Excel.decodeBytes(bytes);
  
  // 3. Select the sheet containing the source data range
  var sheet = excel['Sales Data'];
  
  // 4. Overwrite or append records in the source data rows
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North Region"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(45000));
  
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("South Region"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(32000));
  
  sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("East Region"));
  sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(51000));

  // 5. Save the workbook bytes. 
  // All Pivot Cache Relationships, Definitions, and Filters are fully preserved!
  excel.save(fileName: 'sales_report_pivot.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Full report (headers, formulas, chart)
// ---------------------------------------------------------------------------
const String fullDemoSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateFullReport() {
  var excel = Excel.createExcel();
  var sheetName = 'Financial Report';
  excel.rename('Sheet1', sheetName);
  var sheet = excel[sheetName];

  // Define headers, borders, colors...
  var headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
  );

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
  
  // Calculate Profit with Excel Formulas:
  sheet.updateCell(CellIndex.indexByString("D2"), FormulaCellValue("B2-C2"));
  
  // Summing Totals:
  sheet.updateCell(CellIndex.indexByString("B8"), FormulaCellValue("SUM(B2:B7)"));

  // Add ColumnChart comparing revenue/profit
  sheet.addChart(ColumnChart(
    title: "Monthly Financial Performance",
    series: [
      ChartSeries(name: "Revenue", categoriesRange: "'Financial Report'!\$A\$2:\$A\$7", valuesRange: "'Financial Report'!\$B\$2:\$B\$7"),
      ChartSeries(name: "Profit", categoriesRange: "'Financial Report'!\$A\$2:\$A\$7", valuesRange: "'Financial Report'!\$D\$2:\$D\$7"),
    ],
    anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
  ));

  excel.save(fileName: 'financial_performance_report.xlsx');
}
''';