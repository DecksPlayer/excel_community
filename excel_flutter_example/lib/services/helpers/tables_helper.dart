import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

import '../../data/table_samples.dart';

/// Demo workbook: a sales table with a totals row and structured-reference
/// formulas, plus a gallery sheet with one small table per built-in style.
Excel buildTablesWorkbook() {
  final excel = Excel.createExcel();

  final sales = excel['Sales'];
  const rows = [
    ('North', 120, 2.5), ('South', 95, 3.0), ('East', 140, 1.75),
    ('West', 80, 2.25), ('Central', 110, 2.0),
  ];
  sales.appendRow([TextCellValue('Region'), TextCellValue('Units'), TextCellValue('Price'), TextCellValue('Revenue')]);
  for (var i = 0; i < rows.length; i++) {
    final (region, units, price) = rows[i];
    sales.appendRow([
      TextCellValue(region),
      IntCellValue(units),
      DoubleCellValue(price),
      FormulaCellValue('B${i + 2}*C${i + 2}'),
    ]);
  }
  final table = sales.addTable('A1:D${rows.length + 2}',
      name: 'Sales',
      style: TableStyle.medium(9),
      showTotalsRow: true,
      columns: const [
        TableColumn('Region', totalsLabel: 'Total'),
        TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
        TableColumn('Price', totalsFunction: TableTotalsFunction.average),
        TableColumn('Revenue', totalsFunction: TableTotalsFunction.sum),
      ]);
  sales.updateCell(CellIndex.indexByString('F1'), TextCellValue('Best region revenue'));
  sales.updateCell(CellIndex.indexByString('G1'),
      FormulaCellValue('MAX(${table.columnReference('Revenue')})'));
  for (var c = 0; c < 7; c++) {
    sales.setColumnWidth(c, c == 5 ? 22 : 13);
  }

  // One table per built-in style, four columns apart.
  final gallery = excel['Style Gallery'];
  for (var i = 0; i < builtInTableStyles.length; i++) {
    final (label, style) = builtInTableStyles[i];
    final top = (i ~/ 4) * 6;
    final left = (i % 4) * 4;
    CellIndex at(int r, int c) => CellIndex.indexByColumnRow(columnIndex: left + c, rowIndex: top + r);
    gallery.updateCell(at(0, 0), TextCellValue(label), cellStyle: CellStyle(bold: true));
    gallery.updateCell(at(1, 0), TextCellValue('Item'));
    gallery.updateCell(at(1, 1), TextCellValue('Qty'));
    gallery.updateCell(at(1, 2), TextCellValue('Price'));
    for (var r = 0; r < 3; r++) {
      gallery.updateCell(at(r + 2, 0), TextCellValue('Item ${r + 1}'));
      gallery.updateCell(at(r + 2, 1), IntCellValue(10 + r * 5));
      gallery.updateCell(at(r + 2, 2), DoubleCellValue(1.5 + r));
    }
    final start = at(1, 0).cellId;
    final end = at(4, 2).cellId;
    gallery.addTable('$start:$end', name: style.name, style: style);
  }
  return excel;
}

Future<String> generateTablesHelper() async {
  final excel = buildTablesWorkbook();
  final summary = 'Sales (table with totals), Style Gallery (${builtInTableStyles.length} styles)';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'tables_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Tables workbook generated successfully!\n'
          'Worksheets: $summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: tables_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Tables Example',
      fileName: 'tables_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Tables workbook saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
