import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

/// Demo workbook: a "Links" sheet with every kind of hyperlink and the
/// sheets those internal links point to.
Excel buildHyperlinksWorkbook() {
  final excel = Excel.createExcel();
  final links = excel['Links'];

  links.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hyperlink Showcase'),
      cellStyle: CellStyle(bold: true, fontSize: 16));
  links.appendRow([TextCellValue('Kind'), TextCellValue('Link'), TextCellValue('Code')]);
  for (var col = 0; col < 3; col++) {
    links.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 1)).cellStyle =
        CellStyle(bold: true, backgroundColorHex: ExcelColor.fromHexString('#DBEAFE'));
  }

  final rows = <(String, Hyperlink, String, String)>[
    ('Web page', Hyperlink.url('https://pub.dev/packages/excel_community', tooltip: 'Open on pub.dev'),
        'excel_community on pub.dev', 'Hyperlink.url(...)'),
    ('E-mail', Hyperlink.email('sales@example.com', subject: 'Q3 report'), 'Contact sales',
        'Hyperlink.email(..., subject:)'),
    ('Page anchor', const Hyperlink(url: 'https://dart.dev/language', location: 'variables'),
        'Dart variables', 'Hyperlink(url:, location:)'),
    ('Cell on another sheet', Hyperlink.cell('Q1 Sales', 'B4', tooltip: 'Q1 total'),
        'Go to Q1 total', "Hyperlink.cell('Q1 Sales', 'B4')"),
    ('Range on another sheet', Hyperlink.cell('Q1 Sales', 'A1:B4'), 'Select Q1 table',
        "Hyperlink.cell('Q1 Sales', 'A1:B4')"),
  ];
  for (var i = 0; i < rows.length; i++) {
    final (kind, link, text, code) = rows[i];
    final row = i + 2;
    links.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row), TextCellValue(kind));
    links.setHyperlink(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row), link, text: text);
    links.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: row), TextCellValue(code));
  }
  links.setColumnWidth(0, 24);
  links.setColumnWidth(1, 32);
  links.setColumnWidth(2, 40);

  final q1 = excel['Q1 Sales'];
  q1.appendRow([TextCellValue('Region'), TextCellValue('Sales')]);
  q1.appendRow([TextCellValue('North'), IntCellValue(1200)]);
  q1.appendRow([TextCellValue('South'), IntCellValue(950)]);
  q1.appendRow([TextCellValue('Total'), FormulaCellValue('SUM(B2:B3)')]);
  q1.setHyperlink(CellIndex.indexByString('D1'), Hyperlink.cell('Links', 'A1'), text: '← Back to links');
  q1.setColumnWidth(3, 18);

  return excel;
}

Future<String> generateHyperlinksHelper() async {
  final excel = buildHyperlinksWorkbook();
  const summary = 'Links (web, e-mail, anchor, cell and range links), Q1 Sales';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'hyperlinks_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Hyperlinks workbook generated successfully!\n'
          'Worksheets: $summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: hyperlinks_example.xlsx\n'
          '💡 Ctrl+click a link in Excel to follow it.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Hyperlinks Example',
      fileName: 'hyperlinks_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Hyperlinks workbook saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
