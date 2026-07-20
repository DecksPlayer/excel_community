import 'dart:convert';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Conditional Formatting Tests', () {
    test('Add conditional formatting rule and verify in-memory structure', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final rule1 = ConditionalFormattingRule.cellIs(
        operator: ConditionalFormattingOperator.greaterThan,
        formulae: ['50'],
        backgroundColor: ExcelColor.fromHexString('FFC7CE'),
        fontColor: ExcelColor.fromHexString('9C0006'),
        bold: true,
      );

      sheet.addConditionalFormattingRule('A1:A10', rule1);

      expect(sheet.conditionalFormattings.length, equals(1));
      expect(sheet.conditionalFormattings.first.sqref, equals('A1:A10'));
      expect(sheet.conditionalFormattings.first.rules.length, equals(1));

      final savedRule = sheet.conditionalFormattings.first.rules.first;
      expect(savedRule.type, equals(ConditionalFormattingType.cellIs));
      expect(savedRule.operator, equals(ConditionalFormattingOperator.greaterThan));
      expect(savedRule.formulae, equals(['50']));
      expect(savedRule.style.backgroundColor?.colorHex, equals('FFC7CE'));
      expect(savedRule.style.fontColor?.colorHex, equals('9C0006'));
      expect(savedRule.style.bold, isTrue);
    });

    test('Encode and Decode workbook with conditional formatting rules', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      // Populate dummy data
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(100));

      final greenFill = ExcelColor.fromHexString('#C6EFCE');
      final greenText = ExcelColor.fromHexString('#006100');

      final redFill = ExcelColor.fromHexString('#FFC7CE');
      final redText = ExcelColor.fromHexString('#9C0006');

      final ruleHigh = ConditionalFormattingRule.cellIs(
        operator: ConditionalFormattingOperator.greaterThan,
        formulae: ['50'],
        backgroundColor: greenFill,
        fontColor: greenText,
        bold: true,
        priority: 1,
      );

      final ruleLow = ConditionalFormattingRule.cellIs(
        operator: ConditionalFormattingOperator.lessThanOrEqual,
        formulae: ['50'],
        backgroundColor: redFill,
        fontColor: redText,
        priority: 2,
      );

      sheet.addConditionalFormatting('A1:A10', [ruleHigh, ruleLow]);

      // Rule for text
      final ruleText = ConditionalFormattingRule.containsText(
        text: 'Aprobado',
        backgroundColor: ExcelColor.fromHexString('#800080'),
        fontColor: ExcelColor.fromHexString('#FFFFFF'),
        italic: true,
      );
      sheet.addConditionalFormattingRule('B1:B20', ruleText);

      final bytes = excel.encode()!;
      expect(bytes, isNotEmpty);

      // Re-open workbook and verify parsed rules
      final reloadedExcel = Excel.decodeBytes(bytes);
      final reloadedSheet = reloadedExcel['Sheet1'];

      expect(reloadedSheet.conditionalFormattings.length, equals(2));

      final group1 = reloadedSheet.conditionalFormattings
          .firstWhere((g) => g.sqref == 'A1:A10');
      expect(group1.rules.length, equals(2));

      final r1 = group1.rules[0];
      expect(r1.type, equals(ConditionalFormattingType.cellIs));
      expect(r1.operator, equals(ConditionalFormattingOperator.greaterThan));
      expect(r1.formulae, equals(['50']));
      expect(r1.style.backgroundColor?.colorHex, equals('FFC6EFCE'));
      expect(r1.style.fontColor?.colorHex, equals('FF006100'));
      expect(r1.style.bold, isTrue);

      final r2 = group1.rules[1];
      expect(r2.type, equals(ConditionalFormattingType.cellIs));
      expect(r2.operator, equals(ConditionalFormattingOperator.lessThanOrEqual));
      expect(r2.formulae, equals(['50']));
      expect(r2.style.backgroundColor?.colorHex, equals('FFFFC7CE'));
      expect(r2.style.fontColor?.colorHex, equals('FF9C0006'));

      final group2 = reloadedSheet.conditionalFormattings
          .firstWhere((g) => g.sqref == 'B1:B20');
      expect(group2.rules.length, equals(1));
      final rText = group2.rules[0];
      expect(rText.text, equals('Aprobado'));
      expect(rText.style.backgroundColor?.colorHex, equals('FF800080'));
      expect(rText.style.fontColor?.colorHex, equals('FFFFFFFF'));
      expect(rText.style.italic, isTrue);
    });

    test('Custom expression and duplicate values rules', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.addConditionalFormattingRule(
        'C1:C10',
        ConditionalFormattingRule.expression(
          formula: 'C1>100',
          backgroundColor: ExcelColor.fromHexString('#FFFF00'),
        ),
      );

      sheet.addConditionalFormattingRule(
        'D1:D10',
        ConditionalFormattingRule.duplicateValues(
          backgroundColor: ExcelColor.fromHexString('#FF0000'),
          fontColor: ExcelColor.fromHexString('#FFFFFF'),
        ),
      );

      final bytes = excel.encode()!;
      final reloadedExcel = Excel.decodeBytes(bytes);
      final reloadedSheet = reloadedExcel['Sheet1'];

      final groupC = reloadedSheet.conditionalFormattings
          .firstWhere((g) => g.sqref == 'C1:C10');
      expect(groupC.rules.first.type, equals(ConditionalFormattingType.expression));
      expect(groupC.rules.first.formulae, equals(['C1>100']));
      expect(groupC.rules.first.style.backgroundColor?.colorHex, equals('FFFFFF00'));

      final groupD = reloadedSheet.conditionalFormattings
          .firstWhere((g) => g.sqref == 'D1:D10');
      expect(groupD.rules.first.type, equals(ConditionalFormattingType.duplicateValues));
      expect(groupD.rules.first.style.backgroundColor?.colorHex, equals('FFFF0000'));
    });
  });
}
