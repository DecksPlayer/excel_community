import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('TabColor Model & Normalization Tests', () {
    test('TabColor.fromHex normalizes 6-char hex with FF alpha prefix', () {
      final tabColor = TabColor.fromHex('#4CAF50');
      expect(tabColor.rgb, equals('FF4CAF50'));
      expect(tabColor.colorHex, equals('FF4CAF50'));
      expect(tabColor.colorHex6, equals('4CAF50'));
    });

    test('TabColor.fromHex without leading #', () {
      final tabColor = TabColor.fromHex('E91E63');
      expect(tabColor.rgb, equals('FFE91E63'));
      expect(tabColor.colorHex, equals('FFE91E63'));
      expect(tabColor.colorHex6, equals('E91E63'));
    });

    test('TabColor.fromHex with 8-char ARGB hex', () {
      final tabColor = TabColor.fromHex('80FF5722');
      expect(tabColor.rgb, equals('80FF5722'));
      expect(tabColor.colorHex, equals('80FF5722'));
      expect(tabColor.colorHex6, equals('FF5722'));
    });

    test('TabColor.fromHex with 3-char shorthand #F00', () {
      final tabColor = TabColor.fromHex('#F00');
      expect(tabColor.rgb, equals('FFFF0000'));
      expect(tabColor.colorHex, equals('FFFF0000'));
      expect(tabColor.colorHex6, equals('FF0000'));
    });

    test('TabColor.fromColor with ExcelColor', () {
      final tabColor = TabColor.fromColor(ExcelColor.green);
      expect(tabColor.color, equals(ExcelColor.green));
      expect(tabColor.colorHex, isNotNull);
      expect(tabColor.colorHex6, isNotNull);
    });

    test('TabColor.fromTheme with optional tint', () {
      final tabColor = TabColor.fromTheme(4, tint: 0.3999);
      expect(tabColor.theme, equals(4));
      expect(tabColor.tint, equals(0.3999));
      expect(tabColor.colorHex, isNull);
      expect(tabColor.colorHex6, isNull);
    });

    test('TabColor toXmlString with rgb', () {
      final tabColor = TabColor.fromHex('#00BCD4');
      expect(tabColor.toXmlString(), equals('<tabColor rgb="FF00BCD4"/>'));
    });

    test('TabColor toXmlString with theme and tint', () {
      final tabColor = TabColor.fromTheme(5, tint: -0.25);
      expect(tabColor.toXmlString(), equals('<tabColor theme="5" tint="-0.25"/>'));
    });

    test('TabColor toXmlString with indexed color', () {
      const tabColor = TabColor(indexed: 64);
      expect(tabColor.toXmlString(), equals('<tabColor indexed="64"/>'));
    });

    test('TabColor toXmlString with auto', () {
      const tabColor = TabColor(auto: true);
      expect(tabColor.toXmlString(), equals('<tabColor auto="1"/>'));
    });

    test('TabColor equality and copyWith', () {
      final c1 = TabColor.fromHex('#4CAF50');
      final c2 = TabColor.fromHex('4CAF50');
      expect(c1, equals(c2));

      final c3 = c1.copyWith(rgb: 'FFFF0000');
      expect(c3.colorHex, equals('FFFF0000'));
      expect(c3, isNot(equals(c1)));
    });
  });

  group('Sheet Tab Color API Tests', () {
    test('Default Sheet has no tab color', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      expect(sheet.hasTabColor, isFalse);
      expect(sheet.tabColor, isNull);
    });

    test('setTabColorHex sets custom tab color', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.setTabColorHex('#9C27B0');
      expect(sheet.hasTabColor, isTrue);
      expect(sheet.tabColor?.colorHex, equals('FF9C27B0'));
      expect(sheet.tabColor?.colorHex6, equals('9C27B0'));
    });

    test('setTabColor sets custom ExcelColor', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.setTabColor(ExcelColor.red);
      expect(sheet.hasTabColor, isTrue);
      expect(sheet.tabColor?.color, equals(ExcelColor.red));
    });

    test('clearTabColor removes tab color', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.setTabColorHex('#FF9800');
      expect(sheet.hasTabColor, isTrue);

      sheet.clearTabColor();
      expect(sheet.hasTabColor, isFalse);
      expect(sheet.tabColor, isNull);
    });

    test('removeTabColor alias removes tab color', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.setTabColorHex('#3F51B5');
      expect(sheet.hasTabColor, isTrue);

      sheet.removeTabColor();
      expect(sheet.hasTabColor, isFalse);
      expect(sheet.tabColor, isNull);
    });

    test('Copying/Cloning sheet preserves tab color', () {
      final excel = Excel.createExcel();
      final sheet1 = excel['Sheet1'];
      sheet1.setTabColorHex('#2196F3');

      excel.copy('Sheet1', 'Sheet1_Copy');
      final sheetCopy = excel['Sheet1_Copy'];

      expect(sheetCopy.hasTabColor, isTrue);
      expect(sheetCopy.tabColor?.colorHex, equals('FF2196F3'));
    });
  });

  group('XLSX Tab Color Roundtrip & XML Schema Compliance', () {
    test('Encodes <sheetPr><tabColor .../></sheetPr> in schema order before dimension and sheetViews', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setTabColorHex('#4CAF50');
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Tab Color Test');

      final bytes = excel.encode()!;
      final archive = ZipDecoder().decodeBytes(bytes);

      final sheetXmlFile = archive.findFile('xl/worksheets/sheet1.xml')!;
      final sheetXml = utf8.decode(sheetXmlFile.content);

      // Verify <tabColor> inside <sheetPr>
      expect(sheetXml, contains('<sheetPr><tabColor rgb="FF4CAF50"/>'));

      // Verify schema sequence: sheetPr must precede dimension, sheetViews, and sheetData
      final sheetPrIndex = sheetXml.indexOf('<sheetPr>');
      final sheetViewsIndex = sheetXml.indexOf('<sheetViews>');
      final sheetDataIndex = sheetXml.indexOf('<sheetData>');

      expect(sheetPrIndex, isNonNegative);
      expect(sheetViewsIndex, isNonNegative);
      expect(sheetDataIndex, isNonNegative);

      expect(sheetPrIndex, lessThan(sheetViewsIndex));
      expect(sheetViewsIndex, lessThan(sheetDataIndex));
    });

    test('Decodes tab color accurately from XLSX bytes', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setTabColorHex('#E91E63');
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Pink Tab');

      final bytes = excel.encode()!;
      final decoded = Excel.decodeBytes(bytes);
      final decodedSheet = decoded.sheets['Sheet1']!;

      expect(decodedSheet.hasTabColor, isTrue);
      expect(decodedSheet.tabColor?.colorHex, equals('FFE91E63'));
      expect(decodedSheet.tabColor?.colorHex6, equals('E91E63'));
    });

    test('Decodes theme and tint tab color from XLSX', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.tabColor = TabColor.fromTheme(4, tint: 0.3999);

      final bytes = excel.encode()!;
      final decoded = Excel.decodeBytes(bytes);
      final decodedSheet = decoded.sheets['Sheet1']!;

      expect(decodedSheet.hasTabColor, isTrue);
      expect(decodedSheet.tabColor?.theme, equals(4));
      expect(decodedSheet.tabColor?.tint, closeTo(0.3999, 0.0001));
    });

    test('Modifying tab color and re-saving updates XML correctly', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setTabColorHex('#FF0000');

      final bytes1 = excel.encode()!;
      final decoded1 = Excel.decodeBytes(bytes1);
      final sheet1 = decoded1.sheets['Sheet1']!;
      expect(sheet1.tabColor?.colorHex, equals('FFFF0000'));

      // Change tab color to green
      sheet1.setTabColorHex('#00FF00');
      final bytes2 = decoded1.encode()!;
      final decoded2 = Excel.decodeBytes(bytes2);
      final sheet2 = decoded2.sheets['Sheet1']!;
      expect(sheet2.tabColor?.colorHex, equals('FF00FF00'));

      // Verify raw XML has updated tabColor, no duplicate
      final archive = ZipDecoder().decodeBytes(bytes2);
      final xml = utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content);
      expect(xml, contains('<tabColor rgb="FF00FF00"/>'));
      expect(xml, isNot(contains('FFFF0000')));
    });

    test('Clearing tab color from existing file removes <tabColor> on re-save', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setTabColorHex('#3F51B5');

      final bytes1 = excel.encode()!;
      final decoded1 = Excel.decodeBytes(bytes1);
      final sheet1 = decoded1.sheets['Sheet1']!;
      expect(sheet1.hasTabColor, isTrue);

      // Clear tab color
      sheet1.clearTabColor();
      final bytes2 = decoded1.encode()!;
      final decoded2 = Excel.decodeBytes(bytes2);
      final sheet2 = decoded2.sheets['Sheet1']!;
      expect(sheet2.hasTabColor, isFalse);
      expect(sheet2.tabColor, isNull);

      // Verify raw XML has no tabColor tag
      final archive = ZipDecoder().decodeBytes(bytes2);
      final xml = utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content);
      expect(xml, isNot(contains('<tabColor')));
    });

    test('Multiple sheets can each have different tab colors', () {
      final excel = Excel.createExcel();
      final sheet1 = excel['Sheet1'];
      sheet1.setTabColorHex('#FF5722'); // Orange
      sheet1.cell(CellIndex.indexByString('A1')).value = TextCellValue('Sheet1');

      final sheet2 = excel['Sales'];
      sheet2.setTabColorHex('#4CAF50'); // Green
      sheet2.cell(CellIndex.indexByString('A1')).value = TextCellValue('Sales');

      final sheet3 = excel['Reports'];
      sheet3.setTabColorHex('#2196F3'); // Blue
      sheet3.cell(CellIndex.indexByString('A1')).value = TextCellValue('Reports');

      final sheet4 = excel['NoColor']; // Default no color
      sheet4.cell(CellIndex.indexByString('A1')).value = TextCellValue('No Color');

      final bytes = excel.encode()!;
      final decoded = Excel.decodeBytes(bytes);

      expect(decoded.sheets['Sheet1']?.tabColor?.colorHex6, equals('FF5722'));
      expect(decoded.sheets['Sales']?.tabColor?.colorHex6, equals('4CAF50'));
      expect(decoded.sheets['Reports']?.tabColor?.colorHex6, equals('2196F3'));
      expect(decoded.sheets['NoColor']?.hasTabColor, isFalse);
    });
  });
}
