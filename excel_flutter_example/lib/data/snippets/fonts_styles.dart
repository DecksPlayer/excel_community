// Fonts and Styles snippets: showcasing custom font families and styles.
library;

const String fontsStylesSnippet = r'''
// ============================================================================
// FONT 1: Arial
// ============================================================================
// Example code to generate an Excel file using the Arial font family:

import 'package:excel_community/excel_community.dart';

void generateArialExcel() {
  var excel = Excel.createExcel();
  var sheet = excel['Arial Sheet'];

  var cellStyle = CellStyle(
    fontFamily: getFontFamily(FontFamily.Arial), // Mapped to 'Arial'
    fontSize: 12,
    bold: true,
  );

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = TextCellValue("Arial Bold Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'arial_font_example.xlsx');
}


// ============================================================================
// FONT 2: Courier New
// ============================================================================
// Example code to generate an Excel file using the Courier New font family:

import 'package:excel_community/excel_community.dart';

void generateCourierExcel() {
  var excel = Excel.createExcel();
  var sheet = excel['Courier Sheet'];

  var cellStyle = CellStyle(
    fontFamily: getFontFamily(FontFamily.Courier_New), // Mapped to 'Courier New'
    fontSize: 14,
    italic: true,
  );

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = TextCellValue("Courier Italic Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'courier_font_example.xlsx');
}


// ============================================================================
// FONT 3: Comic Sans MS
// ============================================================================
// Example code to generate an Excel file using the Comic Sans MS font family:

import 'package:excel_community/excel_community.dart';

void generateComicSansExcel() {
  var excel = Excel.createExcel();
  var sheet = excel['Comic Sans Sheet'];

  var cellStyle = CellStyle(
    fontFamily: getFontFamily(FontFamily.Comic_Sans_MS), // Mapped to 'Comic Sans MS'
    fontSize: 11,
  );

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = TextCellValue("Comic Sans Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'comic_sans_font_example.xlsx');
}


// ============================================================================
// FONT 4: Calibri
// ============================================================================
// Example code to generate an Excel file using the Calibri font family:

import 'package:excel_community/excel_community.dart';

void generateCalibriExcel() {
  var excel = Excel.createExcel();
  var sheet = excel['Calibri Sheet'];

  var cellStyle = CellStyle(
    fontFamily: 'Calibri', // Standard string name direct usage
    fontSize: 11,
    fontColorHex: ExcelColor.blue,
  );

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = TextCellValue("Calibri Blue Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'calibri_font_example.xlsx');
}
''';
