// Fonts and Styles snippets: showcasing custom font families and styles.
library;

const String fontsStylesSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Fonts & Styles demo workbook showcasing:
/// - Custom font families with getFontFamily(FontFamily.xyz)
/// - Bold, italic, underline, double underline, and strikethrough
/// - Background colors, font colors, and cell borders
void generateFontsStylesWorkbook() {
  var excel = Excel.createExcel();
  var sheet = excel['Fonts & Styles Demo'];
  excel.delete('Sheet1');

  // Title Block
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('FONTS & STYLES DEMO'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.indigo,
      horizontalAlign: HorizontalAlign.Center,
    ),
  );
  sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('E1'));

  int row = 3;

  final groupHeaderStyle = CellStyle(
    bold: true,
    fontSize: 14,
    backgroundColorHex: ExcelColor.grey200,
    fontColorHex: ExcelColor.black,
  );

  // SECTION 1: FONT FAMILIES
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('1. FONT FAMILIES'),
    cellStyle: groupHeaderStyle,
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('E$row'));
  row += 2;

  final tableHeaderStyle = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.indigo300,
    fontColorHex: ExcelColor.white,
    horizontalAlign: HorizontalAlign.Center,
  );

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('Font Family'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Enum / Mapping'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Sample Text'), cellStyle: tableHeaderStyle);
  row++;

  final fonts = [
    ('Calibri', 'FontFamily.Calibri', getFontFamily(FontFamily.Calibri)),
    ('Arial', 'FontFamily.Arial', getFontFamily(FontFamily.Arial)),
    ('Comic Sans MS', 'FontFamily.Comic_Sans_MS', getFontFamily(FontFamily.Comic_Sans_MS)),
    ('Consolas', 'FontFamily.Consolas', getFontFamily(FontFamily.Consolas)),
    ('Courier New', 'FontFamily.Courier_New', getFontFamily(FontFamily.Courier_New)),
    ('Georgia', 'FontFamily.Georgia', getFontFamily(FontFamily.Georgia)),
    ('Impact', 'FontFamily.Impact', getFontFamily(FontFamily.Impact)),
    ('Times New Roman', 'FontFamily.Times_New_Roman', getFontFamily(FontFamily.Times_New_Roman)),
  ];

  for (final (name, enumName, family) in fonts) {
    sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue(name));
    sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue(enumName));
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue('The quick brown fox jumps over the lazy dog'),
      cellStyle: CellStyle(fontFamily: family, fontSize: 11),
    );
    row++;
  }

  row += 2;

  // SECTION 2: TEXT DECORATION & STYLES
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('2. STYLES & DECORATIONS'),
    cellStyle: groupHeaderStyle,
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('E$row'));
  row += 2;

  final styles = [
    ('Bold', CellStyle(bold: true)),
    ('Italic', CellStyle(italic: true)),
    ('Single Underline', CellStyle(underline: Underline.Single)),
    ('Double Underline', CellStyle(underline: Underline.Double)),
    ('Strikethrough', CellStyle(strikethrough: true)),
    ('Custom Color & Fill', CellStyle(
      bold: true,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.indigo,
    )),
    ('Thin Border', CellStyle(
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
    )),
  ];

  for (final (label, style) in styles) {
    sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue(label));
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue('Sample styled text'),
      cellStyle: style,
    );
    row++;
  }

  // Adjust column widths
  sheet.setColumnWidth(0, 20);
  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 45);

  final bytes = excel.save(fileName: 'fonts_styles_example.xlsx');
  // Or: File('fonts_styles_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
