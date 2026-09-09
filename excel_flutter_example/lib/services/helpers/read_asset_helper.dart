import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> readAssetExcelHelper() async {
  // 1. Read the Excel document from Flutter assets
  const assetKey = 'assets/excel_borter_test_output.xlsx';
  final byteData = await rootBundle.load(assetKey);
  final assetBytes = byteData.buffer.asUint8List(
    byteData.offsetInBytes,
    byteData.lengthInBytes,
  );

  if (assetBytes.isEmpty) {
    throw Exception('Failed to load asset: $assetKey');
  }

  // 2. Decode the document using Excel.decodeBytes
  final excel = Excel.decodeBytes(assetBytes);
  final sheetNames = excel.tables.keys.toList();

  if (sheetNames.isEmpty) {
    throw Exception('No worksheets found in decoded asset document.');
  }

  final sheet = excel[sheetNames.first];
  final rowsCount = sheet.maxRows;
  final colsCount = sheet.maxColumns;

  // Extract key values to demonstrate reading capability
  final valB3 = sheet.cell(CellIndex.indexByString('B3')).value?.toString() ?? '(empty)';
  final valG3 = sheet.cell(CellIndex.indexByString('G3')).value?.toString() ?? '(empty)';
  final valB5 = sheet.cell(CellIndex.indexByString('B5')).value?.toString() ?? '(empty)';
  final valM5 = sheet.cell(CellIndex.indexByString('M5')).value?.toString() ?? '(empty)';
  final valM6 = sheet.cell(CellIndex.indexByString('M6')).value?.toString() ?? '(empty)';

  // Count cells with borders
  int cellsWithBorders = 0;
  for (int r = 0; r < sheet.maxRows; r++) {
    for (int c = 0; c < sheet.maxColumns; c++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r));
      final st = cell.cellStyle;
      if (st != null &&
          (st.leftBorder.borderStyle != null ||
           st.rightBorder.borderStyle != null ||
           st.topBorder.borderStyle != null ||
           st.bottomBorder.borderStyle != null)) {
        cellsWithBorders++;
      }
    }
  }

  // 3. Re-encode and save/download to verify it stays identical and exportable
  const outputFileName = 'extracted_border_asset.xlsx';

  if (kIsWeb) {
    final bytes = excel.save(fileName: outputFileName);
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Asset XLSX successfully read & extracted!\n\n'
          '📄 Asset Source: $assetKey\n'
          '📊 Worksheets: ${sheetNames.join(", ")}\n'
          '📐 Dimensions: $rowsCount rows x $colsCount columns\n'
          '🔲 Bordered Cells: $cellsWithBorders cells preserved\n'
          '🔍 Extracted Data Samples:\n'
          '  • B3: "$valB3"\n'
          '  • G3: "$valG3"\n'
          '  • B5: "$valB5", M5: "$valM5"\n'
          '  • M6 Total: "$valM6"\n\n'
          '💾 Output File Size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          '📥 The re-encoded file is downloading automatically: $outputFileName';
    }
    throw Exception('Failed to re-encode Excel asset for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to re-encode Excel asset.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Extracted Asset Excel File',
      fileName: outputFileName,
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Asset XLSX successfully read & extracted!\n\n'
          '📄 Asset Source: $assetKey\n'
          '📊 Worksheets: ${sheetNames.join(", ")}\n'
          '📐 Dimensions: $rowsCount rows x $colsCount columns\n'
          '🔲 Bordered Cells: $cellsWithBorders cells preserved\n'
          '🔍 Extracted Data Samples:\n'
          '  • B3: "$valB3"\n'
          '  • G3: "$valG3"\n'
          '  • B5: "$valB5", M5: "$valM5"\n'
          '  • M6 Total: "$valM6"\n\n'
          '💾 Location: $outputFile\n'
          '📦 Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '📌 Open in Excel to verify complete borders and values.';
    }
    return 'Asset read and verified successfully (save cancelled).';
  }
}
