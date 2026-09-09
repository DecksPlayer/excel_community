import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

/// Loads a bundled fixture workbook whose formula cells carry a real,
/// pre-calculated `<v>` result (the kind Microsoft Excel writes, but that
/// excel_community itself never produces since it doesn't evaluate
/// formulas), then reads [FormulaCellValue.cachedValue] and [Data.displayText]
/// off of it.
Future<String> generateFormulasDisplayTextHelper() async {
  const assetKey = 'assets/formulas_display_text_example.xlsx';
  final byteData = await rootBundle.load(assetKey);
  final assetBytes = byteData.buffer.asUint8List(
    byteData.offsetInBytes,
    byteData.lengthInBytes,
  );

  if (assetBytes.isEmpty) {
    throw Exception('Failed to load asset: $assetKey');
  }

  final excel = Excel.decodeBytes(assetBytes);
  final sheet = excel['Formulas & Display'];

  // --- 1. FormulaCellValue.cachedValue --------------------------------------
  final formulaRefs = ['D5', 'D6', 'D7', 'D8'];
  final formulaLines = <String>[];
  for (final ref in formulaRefs) {
    final cell = sheet.cell(CellIndex.indexByString(ref));
    final value = cell.value;
    if (value is FormulaCellValue) {
      formulaLines.add(
        '  $ref: "${value.formula}" -> cachedValue=${value.cachedValue} -> displayText="${cell.displayText}"',
      );
    }
  }

  // --- 2. Data.displayText by number format ---------------------------------
  final displayRefs = ['B12', 'B13', 'B14', 'B15', 'B16'];
  final displayLines = <String>[];
  for (final ref in displayRefs) {
    final cell = sheet.cell(CellIndex.indexByString(ref));
    final formatLabel = sheet
        .cell(CellIndex.indexByString('A${cell.rowIndex + 1}'))
        .value
        ?.toString();
    displayLines.add(
      '  $ref ($formatLabel): raw=${cell.value} -> displayText="${cell.displayText}"',
    );
  }

  final summary =
      '📐 FormulaCellValue.cachedValue (read from the fixture):\n'
      '${formulaLines.join('\n')}\n\n'
      '🎨 Data.displayText (rendered from CellStyle.numberFormat):\n'
      '${displayLines.join('\n')}\n\n'
      'ℹ️ Re-encoding this workbook resets FormulaCellValue.cachedValue to null '
      'for every formula cell — excel_community only reads a cached <v>, it '
      'never recalculates or rewrites one.';

  // --- 3. Re-encode & save, same UX as the other asset-reading example -----
  const outputFileName = 'formulas_display_text_roundtrip.xlsx';

  if (kIsWeb) {
    final bytes = excel.save(fileName: outputFileName);
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Formulas & Display Text example loaded!\n\n'
          '$summary\n\n'
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
      dialogTitle: 'Save Formulas & Display Text Example',
      fileName: outputFileName,
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Formulas & Display Text example loaded!\n\n'
          '$summary\n\n'
          '💾 Location: $outputFile\n'
          '📦 Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return '✅ Formulas & Display Text example loaded!\n\n'
        '$summary\n\n'
        '(Save cancelled.)';
  }
}
