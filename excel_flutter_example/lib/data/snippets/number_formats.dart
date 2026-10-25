// Number formats snippets: built-in standard and custom number formatting rules.
library;

const String numberFormatsFullSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Number Formats demo workbook:
/// - One worksheet per format category (Numbers, Currency, Dates, Percentages, etc.)
/// - Demonstrates built-in ECMA-376 formats (0-44) and custom formats
/// - Side-by-side comparison of raw sample, formatted cell, and format code
void generateNumberFormatsWorkbook() {
  final excel = Excel.createExcel();

  final border = Border(borderStyle: BorderStyle.Thin);
  final cellStyle = CellStyle(
    leftBorder: border,
    rightBorder: border,
    topBorder: border,
    bottomBorder: border,
  );
  final headerStyle = cellStyle.copyWith(
    boldVal: true,
    backgroundColorHexVal: ExcelColor.fromHexString('#0D9488'),
    fontColorHexVal: ExcelColor.white,
    horizontalAlignVal: HorizontalAlign.Center,
  );

  // 1. Numbers Sheet
  final numbersSheet = excel['Numbers'];
  numbersSheet.appendRow([
    TextCellValue('Format Name'),
    TextCellValue('Format Code'),
    TextCellValue('Raw Value'),
    TextCellValue('Formatted Cell'),
  ]);
  for (var col = 0; col < 4; col++) {
    numbersSheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0)).cellStyle = headerStyle;
  }

  final numberSamples = [
    ('Integer', NumFormat.standard_1, DoubleCellValue(1234.0)),
    ('2 Decimals', NumFormat.standard_2, DoubleCellValue(1234.56)),
    ('Thousands Separator', NumFormat.standard_3, DoubleCellValue(1234567.0)),
    ('Thousands + 2 Dec', NumFormat.standard_4, DoubleCellValue(1234567.89)),
    ('Percentage', NumFormat.standard_9, DoubleCellValue(0.75)),
    ('Percentage 2 Dec', NumFormat.standard_10, DoubleCellValue(0.1234)),
    ('Scientific / Exponential', NumFormat.standard_11, DoubleCellValue(123456789.0)),
  ];

  for (final (name, format, val) in numberSamples) {
    numbersSheet.appendRow([
      TextCellValue(name),
      TextCellValue(format.formatCode),
      val,
      val,
    ]);
    final lastRow = numbersSheet.maxRows - 1;
    // Apply number format to column D
    numbersSheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: lastRow)).cellStyle =
        cellStyle.copyWith(numberFormat: format, horizontalAlignVal: HorizontalAlign.Right);
  }

  // 2. Currency & Accounting Sheet
  final currencySheet = excel['Currency'];
  currencySheet.appendRow([
    TextCellValue('Format Name'),
    TextCellValue('Format Code'),
    TextCellValue('Raw Value'),
    TextCellValue('Formatted Cell'),
  ]);
  for (var col = 0; col < 4; col++) {
    currencySheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0)).cellStyle = headerStyle;
  }

  final currencySamples = [
    ('Currency $', NumFormat.standard_5, DoubleCellValue(1234.0)),
    ('Currency $ with Cents', NumFormat.standard_7, DoubleCellValue(1234.56)),
    ('Accounting Indent (ID 44)', NumFormat.standard_44, DoubleCellValue(5432.10)),
    ('Custom Euro pattern', const CustomNumericNumFormat('#,##0.00 "€"'), DoubleCellValue(890.50)),
  ];

  for (final (name, format, val) in currencySamples) {
    currencySheet.appendRow([
      TextCellValue(name),
      TextCellValue(format.formatCode),
      val,
      val,
    ]);
    final lastRow = currencySheet.maxRows - 1;
    currencySheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: lastRow)).cellStyle =
        cellStyle.copyWith(numberFormat: format, horizontalAlignVal: HorizontalAlign.Right);
  }

  // 3. Dates & Times Sheet
  final dateSheet = excel['Dates & Times'];
  dateSheet.appendRow([
    TextCellValue('Format Name'),
    TextCellValue('Format Code'),
    TextCellValue('Formatted Cell'),
  ]);
  for (var col = 0; col < 3; col++) {
    dateSheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0)).cellStyle = headerStyle;
  }

  final dateSamples = [
    ('Short Date (m/d/yy)', NumFormat.standard_14, DateCellValue(year: 2026, month: 10, day: 15)),
    ('Day-Month-Year', NumFormat.standard_15, DateCellValue(year: 2026, month: 10, day: 15)),
    ('Time (h:mm)', NumFormat.standard_20, TimeCellValue(hour: 14, minute: 30, second: 0)),
    ('Time (h:mm:ss)', NumFormat.standard_21, TimeCellValue(hour: 14, minute: 30, second: 45)),
  ];

  for (final (name, format, val) in dateSamples) {
    dateSheet.appendRow([
      TextCellValue(name),
      TextCellValue(format.formatCode),
      val,
    ]);
    final lastRow = dateSheet.maxRows - 1;
    dateSheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: lastRow)).cellStyle =
        cellStyle.copyWith(numberFormat: format, horizontalAlignVal: HorizontalAlign.Right);
  }

  // Adjust column widths for all sheets
  for (final name in ['Numbers', 'Currency', 'Dates & Times']) {
    excel[name].setColumnWidth(0, 26);
    excel[name].setColumnWidth(1, 24);
    excel[name].setColumnWidth(2, 20);
    excel[name].setColumnWidth(3, 22);
  }

  // Save the workbook
  final bytes = excel.save(fileName: 'number_formats_example.xlsx');
  // Or: File('number_formats_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
