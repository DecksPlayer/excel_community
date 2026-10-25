const String hyperlinksSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Hyperlinks demo workbook:
/// - External links: Web URLs and Mailto (email with subject)
/// - Internal links: cell reference ('Q1 Sales'!B4) and range selections
/// - Destination worksheet with return link ("<- Back to links")
void generateHyperlinksWorkbook() {
  final excel = Excel.createExcel();

  // 1. "Links" showcase sheet
  final links = excel['Links'];

  links.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Hyperlink Showcase'),
    cellStyle: CellStyle(bold: true, fontSize: 16),
  );

  links.appendRow([
    TextCellValue('Kind'),
    TextCellValue('Link'),
    TextCellValue('Code Snippet'),
  ]);
  for (var col = 0; col < 3; col++) {
    links.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 1)).cellStyle =
        CellStyle(bold: true, backgroundColorHex: ExcelColor.fromHexString('#DBEAFE'));
  }

  final rows = <(String, Hyperlink, String, String)>[
    (
      'Web page',
      Hyperlink.url('https://pub.dev/packages/excel_community', tooltip: 'Open on pub.dev'),
      'excel_community on pub.dev',
      "Hyperlink.url('https://pub.dev/packages/excel_community')",
    ),
    (
      'E-mail',
      Hyperlink.email('sales@example.com', subject: 'Inquiry Q3 report'),
      'Contact sales',
      "Hyperlink.email('sales@example.com', subject: 'Inquiry Q3 report')",
    ),
    (
      'Page anchor',
      const Hyperlink(url: 'https://dart.dev/language', location: 'variables'),
      'Dart variables documentation',
      "Hyperlink(url: 'https://dart.dev/language', location: 'variables')",
    ),
    (
      'Cell on another sheet',
      Hyperlink.cell('Q1 Sales', 'B4', tooltip: 'Jump to Q1 total'),
      'Go to Q1 total',
      "Hyperlink.cell('Q1 Sales', 'B4')",
    ),
    (
      'Range on another sheet',
      Hyperlink.cell('Q1 Sales', 'A1:B4', tooltip: 'Select Q1 data block'),
      'Select Q1 table',
      "Hyperlink.cell('Q1 Sales', 'A1:B4')",
    ),
  ];

  for (var i = 0; i < rows.length; i++) {
    final (kind, link, text, code) = rows[i];
    final row = i + 2;
    links.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row),
      TextCellValue(kind),
    );
    // Set hyperlink with custom display text
    links.setHyperlink(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row),
      link,
      text: text,
    );
    links.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: row),
      TextCellValue(code),
    );
  }

  links.setColumnWidth(0, 24);
  links.setColumnWidth(1, 32);
  links.setColumnWidth(2, 45);

  // 2. Destination sheet "Q1 Sales"
  final q1 = excel['Q1 Sales'];
  q1.appendRow([TextCellValue('Region'), TextCellValue('Sales')]);
  q1.appendRow([TextCellValue('North'), IntCellValue(1200)]);
  q1.appendRow([TextCellValue('South'), IntCellValue(950)]);
  q1.appendRow([TextCellValue('Total'), FormulaCellValue('SUM(B2:B3)')]);

  // Back link jumping back to 'Links'!A1
  q1.setHyperlink(
    CellIndex.indexByString('D1'),
    Hyperlink.cell('Links', 'A1'),
    text: '← Back to links',
  );
  q1.setColumnWidth(3, 18);

  // 3. Save output
  final bytes = excel.save(fileName: 'hyperlinks_example.xlsx');
  // Or: File('hyperlinks_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
