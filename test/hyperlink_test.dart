import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

String _part(List<int> bytes, String name) =>
    utf8.decode(ZipDecoder().decodeBytes(bytes).findFile(name)!.content);

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Hyperlink model', () {
    test('factories build external and internal targets', () {
      expect(Hyperlink.url('https://pub.dev').isExternal, isTrue);
      expect(Hyperlink.email('ana@example.com', subject: 'Q3 report').url,
          'mailto:ana@example.com?subject=Q3%20report');
      expect(Hyperlink.cell('Summary', 'B4').location, 'Summary!B4');
      expect(Hyperlink.cell('Q1 Sales', 'A1').location, "'Q1 Sales'!A1");
      expect(Hyperlink.cell("Bob's", 'A1').location, "'Bob''s'!A1");
      expect(Hyperlink.cell('AB12', 'A1').location, "'AB12'!A1",
          reason: 'names that look like cell references are quoted');
      expect(Hyperlink.location('TotalSales').isExternal, isFalse);
    });

    test('default text', () {
      expect(Hyperlink.url('https://pub.dev').defaultText, 'https://pub.dev');
      expect(Hyperlink.email('a@b.com', subject: 'x').defaultText, 'a@b.com');
      expect(Hyperlink.cell('S', 'A1', display: 'Go').defaultText, 'Go');
    });
  });

  group('Sheet API', () {
    test('setHyperlink writes default text and the hyperlink style', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.setHyperlink(_at('A1'), Hyperlink.url('https://pub.dev'));
      final cell = sheet.cell(_at('A1'));
      expect(cell.value, TextCellValue('https://pub.dev'));
      expect(cell.cellStyle!.underline, Underline.Single);
      expect(cell.cellStyle!.fontColor.colorHex, contains('0563C1'));
      expect(cell.hyperlink, Hyperlink.url('https://pub.dev'));
    });

    test('keeps an existing value unless text is given; styled can be off', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.cell(_at('A1')).value = TextCellValue('Docs');
      sheet.setHyperlink(_at('A1'), Hyperlink.url('https://dart.dev'), styled: false);
      sheet.setHyperlink(_at('A2'), Hyperlink.url('https://dart.dev'), text: 'Dart');
      expect(sheet.cell(_at('A1')).value, TextCellValue('Docs'));
      expect(sheet.cell(_at('A1')).cellStyle?.underline ?? Underline.None, Underline.None);
      expect(sheet.cell(_at('A2')).value, TextCellValue('Dart'));
    });

    test('ranges, lookup and removal', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.setHyperlinkRange('B2:C3', Hyperlink.location('Totals'));
      expect(sheet.getHyperlink(_at('C3'))?.location, 'Totals');
      expect(sheet.getHyperlink(_at('D3')), isNull);

      // A single-cell link inside the range replaces the range.
      sheet.setHyperlink(_at('B2'), Hyperlink.url('https://x.dev'));
      expect(sheet.hyperlinks.keys, ['B2']);

      sheet.cell(_at('B2')).hyperlink = null;
      expect(sheet.hasHyperlinks, isFalse);
      expect(sheet.cell(_at('B2')).value, isNotNull, reason: 'value is kept');
    });

    test('links move with inserted and removed rows and columns', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.setHyperlink(_at('B2'), Hyperlink.url('https://a.dev'));
      sheet.setHyperlinkRange('A5:C6', Hyperlink.location('Range'));
      sheet.setHyperlink(_at('D9'), Hyperlink.url('https://gone.dev'));

      sheet.insertRow(0);
      expect(sheet.hyperlinks.keys, containsAll(['B3', 'A6:C7', 'D10']));
      sheet.insertColumn(0);
      expect(sheet.hyperlinks.keys, containsAll(['C3', 'B6:D7', 'E10']));
      sheet.removeRow(9); // row 10 holds the E10 link
      expect(sheet.hyperlinks.keys.toSet(), {'C3', 'B6:D7'});
      sheet.removeColumn(1); // column B: range shrinks to B6:C7
      expect(sheet.hyperlinks.keys.toSet(), {'B3', 'B6:C7'});
    });

    test('copying a sheet copies its links', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setHyperlink(_at('A1'), Hyperlink.url('https://pub.dev'));
      excel.copy('Sheet1', 'Copy');
      expect(excel['Copy'].getHyperlink(_at('A1')), Hyperlink.url('https://pub.dev'));
    });
  });

  group('XLSX round-trip', () {
    test('external links become relationships with TargetMode External', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setHyperlink(_at('A1'), Hyperlink.url('https://pub.dev/packages?q=excel&sort=top',
          tooltip: 'Search "excel"'));
      sheet.setHyperlink(_at('A2'), Hyperlink.cell('Data Sheet', 'B4'));
      final bytes = excel.encode()!;

      final xml = _part(bytes, 'xl/worksheets/sheet1.xml');
      final rels = _part(bytes, 'xl/worksheets/_rels/sheet1.xml.rels');
      final id = RegExp(r'<hyperlink ref="A1" r:id="(rId\d+)"').firstMatch(xml)!.group(1);
      expect(rels, contains('Id="$id"'));
      expect(rels, contains('Target="https://pub.dev/packages?q=excel&amp;sort=top"'));
      expect(rels, contains('TargetMode="External"'));
      expect(xml, contains('tooltip="Search &quot;excel&quot;"'));
      expect(xml, contains('<hyperlink ref="A2" location="&apos;Data Sheet&apos;!B4"/>'));

      // Schema order: ... conditionalFormatting, dataValidations, hyperlinks, printOptions, pageMargins ...
      expect(xml.indexOf('<hyperlinks>'), lessThan(xml.indexOf('<pageMargins')));
      expect(xml.indexOf('</sheetData>'), lessThan(xml.indexOf('<hyperlinks>')));
    });

    test('decodes what it encodes', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setHyperlink(_at('A1'), Hyperlink.url('https://pub.dev', tooltip: 'pub'));
      sheet.setHyperlink(_at('A2'), Hyperlink.email('ana@example.com'));
      sheet.setHyperlinkRange('B1:C2', Hyperlink.cell('Sheet1', 'Z100', display: 'Far'));
      sheet.setHyperlink(_at('A3'), Hyperlink(url: 'https://dart.dev/docs', location: 'intro'));

      final decoded = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(decoded.hyperlinks, sheet.hyperlinks);
    });

    test('re-saving keeps other relationships and replaces old links', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setHyperlink(_at('A1'), Hyperlink.url('https://old.dev'));
      sheet.cell(_at('B1')).comment = 'keep me';

      final reopened = Excel.decodeBytes(excel.encode()!);
      reopened['Sheet1'].setHyperlink(_at('A1'), Hyperlink.url('https://new.dev'));
      final bytes = reopened.encode()!;
      final rels = _part(bytes, 'xl/worksheets/_rels/sheet1.xml.rels');
      expect(rels, isNot(contains('old.dev')));
      expect(rels, contains('https://new.dev'));
      expect(rels, contains('/comments"'));
      expect(RegExp('Id="(rId\\d+)"').allMatches(rels).map((m) => m.group(1)).toSet().length,
          RegExp('<Relationship ').allMatches(rels).length,
          reason: 'relationship ids are unique');

      final cleared = Excel.decodeBytes(bytes);
      cleared['Sheet1'].clearHyperlinks();
      final finalBytes = cleared.encode()!;
      expect(_part(finalBytes, 'xl/worksheets/_rels/sheet1.xml.rels'), isNot(contains('/hyperlink"')));
      expect(_part(finalBytes, 'xl/worksheets/sheet1.xml'), isNot(contains('<hyperlinks>')));
    });

    test('links on a new sheet are saved with its own relationships', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(_at('A1')).value = TextCellValue('x');
      excel['Links'].setHyperlink(_at('A1'), Hyperlink.url('https://pub.dev'));
      final decoded = Excel.decodeBytes(excel.encode()!);
      expect(decoded['Links'].getHyperlink(_at('A1'))?.url, 'https://pub.dev');
    });
  });
}
