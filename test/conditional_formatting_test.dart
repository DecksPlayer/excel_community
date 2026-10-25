import 'dart:convert';
import 'package:archive/archive.dart';
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

  group('styles of a reopened file with conditional formatting', () {
    XmlDocument stylesOf(List<int> xlsx) => XmlDocument.parse(utf8.decode(
        ZipDecoder().decodeBytes(xlsx).findFile('xl/styles.xml')!.content));

    List<int> withStyles(List<int> xlsx, String Function(String) edit) {
      final out = Archive();
      for (final f in ZipDecoder().decodeBytes(xlsx).files.where((f) => f.isFile)) {
        var content = f.content as List<int>;
        if (f.name == 'xl/styles.xml') content = utf8.encode(edit(utf8.decode(content)));
        out.addFile(ArchiveFile(f.name, content.length, content));
      }
      return ZipEncoder().encode(out);
    }

    /// Asserts every cellXfs id points to an existing font, fill and border.
    void expectValidIds(XmlDocument styles) {
      final root = styles.rootElement;
      int count(String container, String child) {
        final el = root.getElement(container)!;
        final n = el.findElements(child).length;
        expect(el.getAttribute('count'), '$n', reason: '<$container count>');
        return n;
      }

      final fonts = count('fonts', 'font');
      final fills = count('fills', 'fill');
      final borders = count('borders', 'border');
      for (final xf in root.getElement('cellXfs')!.findElements('xf')) {
        expect(int.parse(xf.getAttribute('fontId')!), lessThan(fonts));
        expect(int.parse(xf.getAttribute('fillId')!), lessThan(fills));
        expect(int.parse(xf.getAttribute('borderId')!), lessThan(borders));
      }
    }

    test('a new style does not count the fonts of <dxfs>', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(5));
      sheet.addConditionalFormattingRule(
          'A1:A10',
          ConditionalFormattingRule.containsText(
              text: 'x',
              backgroundColor: ExcelColor.fromHexString('#800080'),
              fontColor: ExcelColor.fromHexString('#FFFFFF'),
              italic: true));

      final reopened = Excel.decodeBytes(excel.encode()!);
      final style = CellStyle(bold: true, fontSize: 20);
      reopened['Sheet1'].updateCell(
          CellIndex.indexByString('D1'), TextCellValue('new'),
          cellStyle: style);
      final bytes = reopened.encode()!;

      expectValidIds(stylesOf(bytes));
      final d1 = Excel.decodeBytes(bytes)['Sheet1']
          .cell(CellIndex.indexByString('D1'))
          .cellStyle!;
      expect(d1.isBold, isTrue);
      expect(d1.fontSize, 20);
    });

    test('a gradient fill keeps the fill ids of new styles in place', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      // A gradient <fill> at the end of <fills>, as Excel writes it.
      final bytes = withStyles(excel.encode()!, (xml) {
        final fills = RegExp(r'<fills count="(\d+)">').firstMatch(xml)!;
        final n = int.parse(fills.group(1)!);
        return xml
            .replaceFirst(fills.group(0)!, '<fills count="${n + 1}">')
            .replaceFirst('</fills>',
                '<fill><gradientFill degree="90"><stop position="0"><color rgb="FF000000"/></stop>'
                    '<stop position="1"><color rgb="FFFFFFFF"/></stop></gradientFill></fill></fills>');
      });

      final reopened = Excel.decodeBytes(bytes);
      final red = ExcelColor.fromHexString('#FF0000');
      reopened['Sheet1'].updateCell(
          CellIndex.indexByString('B1'), TextCellValue('red'),
          cellStyle: CellStyle(backgroundColorHex: red));
      final saved = reopened.encode()!;

      final styles = stylesOf(saved);
      expectValidIds(styles);
      final fills = styles.rootElement.getElement('fills')!.findElements('fill').toList();
      final xfs = styles.rootElement.getElement('cellXfs')!.findElements('xf').toList();
      final s = RegExp(r'<c r="B1" s="(\d+)"').firstMatch(utf8.decode(
          ZipDecoder().decodeBytes(saved).findFile('xl/worksheets/sheet1.xml')!.content))!;
      final fill = fills[int.parse(xfs[int.parse(s.group(1)!)].getAttribute('fillId')!)];
      expect(fill.findAllElements('fgColor').single.getAttribute('rgb'), 'FFFF0000');
    });
  });
}
