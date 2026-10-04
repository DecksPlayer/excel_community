import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

String _sheetXml(List<int> bytes, [String name = 'xl/worksheets/sheet1.xml']) =>
    utf8.decode(ZipDecoder().decodeBytes(bytes).findFile(name)!.content);

void main() {
  group('OutlineSettings', () {
    test('defaults are Excel\'s and write nothing', () {
      expect(const OutlineSettings().isDefault, isTrue);
      expect(const OutlineSettings().toXmlString(), isEmpty);
      expect(const OutlineSettings(summaryBelow: false, showOutlineSymbols: false).toXmlString(),
          '<outlinePr summaryBelow="0" showOutlineSymbols="0"/>');
    });

    test('new workbooks start with the defaults on every sheet', () {
      final excel = Excel.createExcel();
      expect(excel['Sheet1'].outlineSettings, const OutlineSettings());
      expect(excel['Other'].outlineSettings, const OutlineSettings());
    });
  });

  group('Row grouping', () {
    test('groups, nests and lists groups', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupRows(1, 6);
      sheet.groupRows(2, 3);
      expect(sheet.getRowOutlineLevel(0), 0);
      expect(sheet.getRowOutlineLevel(2), 2);
      expect(sheet.rowGroups, const [OutlineGroup(1, 6, 1), OutlineGroup(2, 3, 2)]);
    });

    test('collapse hides rows and marks the summary row', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupRows(1, 4, collapsed: true);
      expect(sheet.getHiddenRows, {1, 2, 3, 4});
      expect(sheet.rowGroups.single.collapsed, isTrue);

      sheet.expandRowGroup(1, 4);
      expect(sheet.getHiddenRows, isEmpty);
      expect(sheet.rowGroups.single.collapsed, isFalse);
    });

    test('expanding an outer group keeps collapsed inner groups', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupRows(1, 8);
      sheet.groupRows(2, 4, collapsed: true);
      sheet.collapseRowGroup(1, 8);
      sheet.expandRowGroup(1, 8);
      expect(sheet.getHiddenRows, {2, 3, 4});
    });

    test('summary above uses the row before the group', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
      sheet.groupRows(3, 5, collapsed: true);
      expect(sheet.rowGroups.single.collapsed, isTrue);
      final xml = _sheetXml(excel.encode()!);
      expect(xml, contains('<row r="3" collapsed="1">'));
      expect(xml, contains('<outlinePr summaryBelow="0"/>'));
    });

    test('ungroup removes one level and expands a collapsed group', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupRows(1, 3);
      sheet.groupRows(1, 3, collapsed: true);
      sheet.ungroupRows(1, 3);
      expect(sheet.getRowOutlineLevel(1), 1);
      expect(sheet.getHiddenRows, isEmpty);
      sheet.ungroupRows(1, 3);
      expect(sheet.rowGroups, isEmpty);
    });

    test('limits', () {
      final sheet = Excel.createExcel()['Sheet1'];
      for (var i = 0; i < 7; i++) {
        sheet.groupRows(0, 1);
      }
      expect(() => sheet.groupRows(0, 1), throwsRangeError);
      expect(() => sheet.groupRows(3, 2), throwsRangeError);
      expect(() => sheet.setRowOutlineLevel(0, 8), throwsRangeError);
    });
  });

  group('Column grouping', () {
    test('group, collapse and expand columns', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupColumns(1, 3, collapsed: true);
      expect(sheet.getColumnOutlineLevel(2), 1);
      expect(sheet.getHiddenColumns, {1, 2, 3});
      expect(sheet.columnGroups.single, const OutlineGroup(1, 3, 1, collapsed: true));
      sheet.expandColumnGroup(1, 3);
      expect(sheet.getHiddenColumns, isEmpty);
    });

    test('clearGrouping shows everything again', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.groupRows(1, 2, collapsed: true);
      sheet.groupColumns(0, 1, collapsed: true);
      sheet.clearGrouping();
      expect(sheet.rowGroups, isEmpty);
      expect(sheet.columnGroups, isEmpty);
      expect(sheet.getHiddenRows, isEmpty);
      expect(sheet.getHiddenColumns, isEmpty);
    });
  });

  group('Row and column inserts move row/column properties', () {
    test('rows: height, hidden and outline follow their row', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.cell(CellIndex.indexByString('A10')).value = TextCellValue('x');
      sheet.setRowHeight(3, 30);
      sheet.groupRows(4, 6, collapsed: true);

      sheet.insertRow(0);
      expect(sheet.getRowHeight(4), 30);
      expect(sheet.getHiddenRows, {5, 6, 7});
      expect(sheet.rowGroups.single, const OutlineGroup(5, 7, 1, collapsed: true));

      sheet.removeRow(0);
      sheet.removeRow(5); // inside the group
      expect(sheet.rowGroups.single, const OutlineGroup(4, 5, 1, collapsed: true));
    });

    test('columns: width, hidden and outline follow their column', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.cell(CellIndex.indexByString('H1')).value = TextCellValue('x');
      sheet.setColumnWidth(2, 25);
      sheet.groupColumns(3, 4);
      sheet.insertColumn(0);
      expect(sheet.getColumnWidth(3), 25);
      expect(sheet.columnGroups.single, const OutlineGroup(4, 5, 1));
    });
  });

  group('XLSX round-trip', () {
    test('writes outline attributes and sheetFormatPr levels', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Region');
      sheet.groupRows(1, 4);
      sheet.groupRows(2, 3, collapsed: true);
      sheet.groupColumns(1, 2, collapsed: true);
      final xml = _sheetXml(excel.encode()!);

      expect(xml, contains('outlineLevelRow="2"'));
      expect(xml, contains('outlineLevelCol="1"'));
      expect(xml, contains('<row r="2" outlineLevel="1">'), reason: 'empty grouped rows are written');
      expect(xml, contains('<row r="3" hidden="1" outlineLevel="2">'));
      expect(xml, contains('<row r="5" outlineLevel="1" collapsed="1">'), reason: 'summary of the inner group');
      expect(xml, matches(RegExp(r'<col min="2" max="2" [^>]*hidden="1" outlineLevel="1"/>')));
      expect(xml, matches(RegExp(r'<col min="4" max="4" [^>]*collapsed="1"/>')));
      expect(xml, isNot(contains('<outlinePr')), reason: 'default settings');
    });

    test('decodes what it encodes', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false, summaryRight: false);
      sheet.groupRows(2, 9);
      sheet.groupRows(3, 5, collapsed: true);
      sheet.groupColumns(2, 4, collapsed: true);

      final decoded = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(decoded.outlineSettings, sheet.outlineSettings);
      expect(decoded.rowGroups, sheet.rowGroups);
      expect(decoded.columnGroups, sheet.columnGroups);
      expect(decoded.getHiddenRows, sheet.getHiddenRows);
      expect(decoded.getHiddenColumns, sheet.getHiddenColumns);
    });

    test('row heights of empty rows are saved', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
      excel['Sheet1'].setRowHeight(5, 42);
      final decoded = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(decoded.getRowHeight(5), 42);
    });
  });
}

