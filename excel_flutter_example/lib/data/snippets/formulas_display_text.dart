// Formulas & Display Text snippets: FormulaCellValue.cachedValue and
// Data.displayText.
library;

const String formulasDisplayTextSnippet = r'''
// ============================================================================
// 1. FormulaCellValue.cachedValue — reading a formula's last computed result
// ============================================================================
// excel_community never recalculates formulas — it only *reads* the pre-
// calculated <v> result Excel already cached in the file, if the file had
// one. That result lives on FormulaCellValue.cachedValue.

import 'package:excel_community/excel_community.dart';

void inspectFormulaCells(Excel excel) {
  final sheet = excel['Formulas & Display'];
  final cell = sheet.cell(CellIndex.indexByString('D8'));

  final value = cell.value;
  if (value is FormulaCellValue) {
    print('formula: ${value.formula}');       // e.g. "SUM(D5:D7)"
    print('cachedValue: ${value.cachedValue}'); // e.g. IntCellValue(123)
    // null if the workbook was created programmatically and never opened
    // in a real spreadsheet app (nothing has calculated a result yet).
  }
}


// ============================================================================
// 2. Data.displayText — rendering a cell the way a spreadsheet app would
// ============================================================================
// displayText formats the cell's value using its assigned CellStyle.numberFormat
// — no need to hand-roll currency/percentage/date string formatting yourself.
// It also understands FormulaCellValue: it renders the cachedValue, not the
// formula text.

void printDisplayText(Excel excel) {
  var sheet = excel['Formulas & Display'];

  var currencyCell = sheet.cell(CellIndex.indexByString('B12'));
  currencyCell.value = DoubleCellValue(1234.5);
  currencyCell.cellStyle = CellStyle(numberFormat: NumFormat.standard_7);
  print(currencyCell.displayText); // "$1,234.50"

  var percentCell = sheet.cell(CellIndex.indexByString('B13'));
  percentCell.value = DoubleCellValue(0.4567);
  percentCell.cellStyle = CellStyle(numberFormat: NumFormat.standard_10);
  print(percentCell.displayText); // "45.67%"

  var dateCell = sheet.cell(CellIndex.indexByString('B14'));
  dateCell.value = DateCellValue(year: 2026, month: 5, day: 31);
  dateCell.cellStyle = CellStyle(numberFormat: NumFormat.standard_14);
  print(dateCell.displayText); // "05-31-26"

  // Custom format codes work too:
  var customCell = sheet.cell(CellIndex.indexByString('B15'));
  customCell.value = DoubleCellValue(999.9);
  customCell.cellStyle = CellStyle(
    numberFormat: NumFormat.custom(formatCode: '#,##0.0"kg"'),
  );
  print(customCell.displayText); // "999.9kg"
}

// Note: this is a best-effort renderer covering the ~50 built-in ECMA-376
// standard formats plus common custom currency/percentage/date/time
// patterns — not a full implementation of the Excel formatting language.
''';
