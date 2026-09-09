// Snippet demonstrating how to read an XLSX file from Flutter Assets,
// extract its tables/cells/borders, inspect or modify the content,
// and re-save or re-encode it.
library;

const String readAssetSnippet = r'''
import 'package:flutter/services.dart' show rootBundle;
import 'package:excel_community/excel_community.dart';

Future<void> readAndExtractAssetExcel() async {
  // 1. Load the XLSX file bundled as a Flutter Asset
  final byteData = await rootBundle.load('assets/excel_borter_test_output.xlsx');
  final bytes = byteData.buffer.asUint8List(
    byteData.offsetInBytes,
    byteData.lengthInBytes,
  );

  // 2. Decode the bytes into an Excel object
  final excel = Excel.decodeBytes(bytes);

  // 3. Inspect sheets and extract data
  print('Worksheet names: ${excel.tables.keys.toList()}');
  final sheet = excel['Sheet1'];
  print('Dimensions: ${sheet.maxRows} rows x ${sheet.maxColumns} columns');

  // 4. Access cell values, formulas and cell border styling
  final headerB3 = sheet.cell(CellIndex.indexByString('B3'));
  print('B3 value: ${headerB3.value}'); // "Merge b:f"
  print('B3 left border: ${headerB3.cellStyle?.leftBorder.borderStyle}'); // BorderStyle.Thin

  final headerG3 = sheet.cell(CellIndex.indexByString('G3'));
  print('G3 value: ${headerG3.value}'); // "merge columns g:k"

  final cellM6 = sheet.cell(CellIndex.indexByString('M6'));
  print('M6 total: ${cellM6.value}'); // "5"

  // 5. Optionally modify or append new data
  sheet.updateCell(
    CellIndex.indexByString('B8'),
    TextCellValue('Extracted & verified from Asset!'),
    cellStyle: CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#107C41'),
      bottomBorder: Border(borderStyle: BorderStyle.Medium),
    ),
  );

  // 6. Save or re-encode the workbook with 100% border & style fidelity
  excel.save(fileName: 'extracted_border_asset.xlsx');
}
''';
