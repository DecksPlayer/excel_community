import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

/// Demo workbook: yearly sales with months grouped per quarter (rows) and
/// regions per quarter grouped by columns, nested up to two levels.
Excel buildGroupingWorkbook() {
  final excel = Excel.createExcel();
  final sheet = excel['Sales 2026'];

  const regions = ['North', 'South', 'East', 'West'];
  const quarters = {
    'Q1': ['Jan', 'Feb', 'Mar'],
    'Q2': ['Apr', 'May', 'Jun'],
    'Q3': ['Jul', 'Aug', 'Sep'],
    'Q4': ['Oct', 'Nov', 'Dec'],
  };

  final header = CellStyle(bold: true, backgroundColorHex: ExcelColor.fromHexString('#FFEDD5'));
  final total = CellStyle(bold: true);
  sheet.appendRow([TextCellValue('Period'), for (final r in regions) TextCellValue(r), TextCellValue('Total')]);
  for (var c = 0; c <= regions.length + 1; c++) {
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0)).cellStyle = header;
  }

  var row = 1;
  final halfStart = <String, int>{};
  final quarterTotals = <int>[];
  for (final entry in quarters.entries) {
    if (entry.key == 'Q1' || entry.key == 'Q3') halfStart[entry.key] = row;
    final first = row;
    for (final month in entry.value) {
      sheet.appendRow([
        TextCellValue(month),
        for (var i = 0; i < regions.length; i++) IntCellValue(100 + (row * 37 + i * 53) % 150),
        FormulaCellValue('SUM(B${row + 1}:E${row + 1})'),
      ]);
      row++;
    }
    // Quarter summary row below its months.
    sheet.appendRow([
      TextCellValue('${entry.key} total'),
      for (final col in ['B', 'C', 'D', 'E', 'F']) FormulaCellValue('SUM($col${first + 1}:$col$row)'),
    ]);
    for (var c = 0; c <= regions.length + 1; c++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: row)).cellStyle = total;
    }
    sheet.groupRows(first, row - 1);
    quarterTotals.add(row);
    row++;
    if (entry.key == 'Q2' || entry.key == 'Q4') {
      // Half-year summary row below its two quarters.
      final start = halfStart[entry.key == 'Q2' ? 'Q1' : 'Q3']!;
      final totals = quarterTotals.sublist(quarterTotals.length - 2);
      sheet.appendRow([
        TextCellValue(entry.key == 'Q2' ? 'H1 total' : 'H2 total'),
        for (final col in ['B', 'C', 'D', 'E', 'F'])
          FormulaCellValue('$col${totals[0] + 1}+$col${totals[1] + 1}'),
      ]);
      sheet.groupRows(start, row - 1);
      row++;
    }
  }

  // Second half starts collapsed; regions are grouped as one column block.
  sheet.collapseRowGroup(halfStart['Q3']!, row - 2);
  sheet.groupColumns(1, regions.length);
  sheet.setColumnWidth(0, 14);
  for (var c = 1; c <= regions.length + 1; c++) {
    sheet.setColumnWidth(c, 12);
  }
  return excel;
}

Future<String> generateGroupingHelper() async {
  final excel = buildGroupingWorkbook();
  const summary = 'Sales 2026 (months in quarters, quarters in halves, regions grouped)';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'grouping_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Grouping workbook generated successfully!\n'
          'Worksheets: $summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: grouping_example.xlsx\n'
          '💡 Use the 1-2-3 buttons and +/- on the left and top in Excel.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Grouping Example',
      fileName: 'grouping_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Grouping workbook saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
