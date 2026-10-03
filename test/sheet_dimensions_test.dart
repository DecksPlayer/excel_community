import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  // Sheets created with excel['Name'] (not the template 'Sheet1') start without
  // <sheetFormatPr> defaults, so _defaultColumnWidth / _defaultRowHeight are null.
  Sheet newSheet() => Excel.createExcel()['New sheet'];

  group('getColumnWidth / getRowHeight', () {
    test('return Excel defaults on a new sheet without throwing', () {
      final sheet = newSheet();

      expect(sheet.defaultColumnWidth, isNull);
      expect(sheet.defaultRowHeight, isNull);
      expect(sheet.getColumnWidth(3), 8.43);
      expect(sheet.getRowHeight(7), 15.0);
    });

    test('return the explicit value when one was set', () {
      final sheet = newSheet();
      sheet.setColumnWidth(2, 30);
      sheet.setRowHeight(4, 42);

      expect(sheet.getColumnWidth(2), 30);
      expect(sheet.getRowHeight(4), 42);
      // Other columns / rows keep the fallback
      expect(sheet.getColumnWidth(0), 8.43);
      expect(sheet.getRowHeight(0), 15.0);
    });

    test('prefer the sheet default over the Excel default', () {
      final sheet = newSheet();
      sheet.setDefaultColumnWidth(12.5);
      sheet.setDefaultRowHeight(20);

      expect(sheet.getColumnWidth(5), 12.5);
      expect(sheet.getRowHeight(5), 20);
    });

    test('keep the template defaults of Sheet1', () {
      final sheet = Excel.createExcel()['Sheet1'];

      expect(sheet.getColumnWidth(0), sheet.defaultColumnWidth);
    });

    test('work after a save/decode round trip', () {
      final excel = Excel.createExcel();
      excel['New sheet']
          .updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      final sheet = Excel.decodeBytes(excel.encode()!)['New sheet'];

      expect(() => sheet.getColumnWidth(10), returnsNormally);
      expect(() => sheet.getRowHeight(10), returnsNormally);
    });
  });
}
