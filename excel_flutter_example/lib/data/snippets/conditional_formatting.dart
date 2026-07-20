const String conditionalFormattingSnippet = '''
import 'package:excel_community/excel_community.dart';

void main() {
  final excel = Excel.createExcel();

  // ===========================================================================
  // SHEET 1: Semester 1 (Numeric Range & Text Match Rules)
  // ===========================================================================
  excel.rename('Sheet1', 'Semester 1');
  final sheet1 = excel['Semester 1'];

  sheet1.appendRow([TextCellValue('Student'), TextCellValue('Grade'), TextCellValue('Status')]);
  sheet1.appendRow([TextCellValue('John'), IntCellValue(85), TextCellValue('Passed')]);
  sheet1.appendRow([TextCellValue('Mary'), IntCellValue(42), TextCellValue('Failed')]);
  sheet1.appendRow([TextCellValue('Charles'), IntCellValue(95), TextCellValue('Passed')]);
  sheet1.appendRow([TextCellValue('Anna'), IntCellValue(38), TextCellValue('Failed')]);

  // Pass (> 70) -> Green Background (#2E7D32) & White Text
  final passRule = ConditionalFormattingRule.cellIs(
    operator: ConditionalFormattingOperator.greaterThan,
    formulae: ['70'],
    backgroundColor: ExcelColor.fromHexString('2E7D32'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );

  // Fail (<= 50) -> Red Background (#D32F2F) & White Text
  final failRule = ConditionalFormattingRule.cellIs(
    operator: ConditionalFormattingOperator.lessThanOrEqual,
    formulae: ['50'],
    backgroundColor: ExcelColor.fromHexString('D32F2F'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );

  sheet1.addConditionalFormatting('B2:B5', [passRule, failRule]);

  // Text match 'Passed' -> Blue Background (#1976D2)
  final passedRule = ConditionalFormattingRule.containsText(
    text: 'Passed',
    backgroundColor: ExcelColor.fromHexString('1976D2'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet1.addConditionalFormattingRule('C2:C5', passedRule);

  // ===========================================================================
  // SHEET 2: Semester 2 (Duplicates & Formula Expression Rules)
  // ===========================================================================
  final sheet2 = excel['Semester 2'];

  sheet2.appendRow([TextCellValue('Subject'), TextCellValue('Score'), TextCellValue('Status')]);
  sheet2.appendRow([TextCellValue('Mathematics'), IntCellValue(92), TextCellValue('Passed')]);
  sheet2.appendRow([TextCellValue('Physics'), IntCellValue(48), TextCellValue('Failed')]);
  sheet2.appendRow([TextCellValue('Mathematics'), IntCellValue(88), TextCellValue('Passed')]); // Duplicate
  sheet2.appendRow([TextCellValue('Chemistry'), IntCellValue(96), TextCellValue('Passed')]);

  // Rule 1: Duplicate values -> Highlight duplicate subjects in Orange
  final duplicateRule = ConditionalFormattingRule.duplicateValues(
    backgroundColor: ExcelColor.fromHexString('E65100'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('A2:A5', duplicateRule);

  // Rule 2: Custom Expression formula (Score > 90) -> Gold Background
  final topScoreRule = ConditionalFormattingRule.expression(
    formula: 'B2>90',
    backgroundColor: ExcelColor.fromHexString('FBC02D'),
    fontColor: ExcelColor.fromHexString('000000'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('B2:B5', topScoreRule);

  // Rule 3: Text match 'Failed' -> Dark Red Background
  final failedTextRule = ConditionalFormattingRule.containsText(
    text: 'Failed',
    backgroundColor: ExcelColor.fromHexString('B71C1C'),
    fontColor: ExcelColor.fromHexString('FFFFFF'),
    bold: true,
  );
  sheet2.addConditionalFormattingRule('C2:C5', failedTextRule);

  // Save Workbook
  final bytes = excel.save();
}
''';
