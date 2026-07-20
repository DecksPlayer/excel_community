import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateConditionalFormattingHelper() async {
  final excel = Excel.createExcel();

  // ===========================================================================
  // SHEET 1: Semester 1 (Numeric Range & Text Match Rules)
  // ===========================================================================
  excel.rename('Sheet1', 'Semester 1');
  final sheet1 = excel['Semester 1'];

  sheet1.appendRow([
    TextCellValue('Student'),
    TextCellValue('Grade'),
    TextCellValue('Status'),
  ]);

  sheet1.appendRow([TextCellValue('John'), IntCellValue(85), TextCellValue('Passed')]);
  sheet1.appendRow([TextCellValue('Mary'), IntCellValue(42), TextCellValue('Failed')]);
  sheet1.appendRow([TextCellValue('Charles'), IntCellValue(95), TextCellValue('Passed')]);
  sheet1.appendRow([TextCellValue('Anna'), IntCellValue(38), TextCellValue('Failed')]);
  sheet1.appendRow([TextCellValue('Louis'), IntCellValue(70), TextCellValue('Passed')]);

  // Sheet 1 Rule 1 & 2: Numeric range (Grade > 70 = Green, Grade <= 50 = Red)
  final passRule = ConditionalFormattingRule.cellIs(
    operator: ConditionalFormattingOperator.greaterThan,
    formulae: ['70'],
    backgroundColor: ExcelColor.fromHexString('2E7D32'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );

  final failRule = ConditionalFormattingRule.cellIs(
    operator: ConditionalFormattingOperator.lessThanOrEqual,
    formulae: ['50'],
    backgroundColor: ExcelColor.fromHexString('D32F2F'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );

  sheet1.addConditionalFormatting('B2:B6', [passRule, failRule]);

  // Sheet 1 Rule 3: Text match 'Passed' -> Blue Background
  final passedRule = ConditionalFormattingRule.containsText(
    text: 'Passed',
    backgroundColor: ExcelColor.fromHexString('1976D2'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet1.addConditionalFormattingRule('C2:C6', passedRule);

  // ===========================================================================
  // SHEET 2: Semester 2 (Duplicate Values & Expression Formula Rules)
  // ===========================================================================
  final sheet2 = excel['Semester 2'];

  sheet2.appendRow([
    TextCellValue('Subject'),
    TextCellValue('Score'),
    TextCellValue('Status'),
  ]);

  sheet2.appendRow([TextCellValue('Mathematics'), IntCellValue(92), TextCellValue('Passed')]);
  sheet2.appendRow([TextCellValue('Physics'), IntCellValue(48), TextCellValue('Failed')]);
  sheet2.appendRow([TextCellValue('Mathematics'), IntCellValue(88), TextCellValue('Passed')]); // Duplicate
  sheet2.appendRow([TextCellValue('Chemistry'), IntCellValue(96), TextCellValue('Passed')]);
  sheet2.appendRow([TextCellValue('History'), IntCellValue(30), TextCellValue('Failed')]);

  // Sheet 2 Rule 1: Duplicate values -> Highlight duplicate subjects in Orange
  final duplicateRule = ConditionalFormattingRule.duplicateValues(
    backgroundColor: ExcelColor.fromHexString('E65100'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('A2:A6', duplicateRule);

  // Sheet 2 Rule 2: Custom Expression formula (Score > 90) -> Gold Background
  final topScoreRule = ConditionalFormattingRule.expression(
    formula: 'B2>90',
    backgroundColor: ExcelColor.fromHexString('FBC02D'),
    fontColor: ExcelColor.fromHexString('000000'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('B2:B6', topScoreRule);

  // Sheet 2 Rule 3: Text match 'Failed' -> Dark Red Background
  final failedTextRule = ConditionalFormattingRule.containsText(
    text: 'Failed',
    backgroundColor: ExcelColor.fromHexString('B71C1C'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('C2:C6', failedTextRule);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'conditional_formatting.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Multi-sheet Excel with Conditional Formatting generated!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 Sheets: Semester 1 & Semester 2';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Multi-sheet Conditional Formatting Excel File',
      fileName: 'conditional_formatting.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Multi-sheet Conditional Formatting Excel saved!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 Open with Excel to inspect multiple formatted sheets!';
    }
    return 'Save cancelled.';
  }
}
