import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('Freeze Panes Tests', () {
    test('Set, save, and decode frozen rows and columns', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Set frozen row/column properties
      sheet.frozenRows = 2;
      sheet.frozenColumns = 3;

      expect(sheet.frozenRows, equals(2));
      expect(sheet.frozenColumns, equals(3));

      // Save to bytes
      var bytes = excel.encode();
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, isTrue);

      // Decode the bytes back to a new Excel workbook
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      // Assert properties are correctly preserved and parsed back
      expect(sheet2.frozenRows, equals(2));
      expect(sheet2.frozenColumns, equals(3));
    });

    test('Freeze rows only preservation', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.frozenRows = 1;
      sheet.frozenColumns = null;

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      expect(sheet2.frozenRows, equals(1));
      expect(sheet2.frozenColumns, isNull);
    });

    test('Freeze columns only preservation', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.frozenRows = null;
      sheet.frozenColumns = 2;

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      expect(sheet2.frozenRows, isNull);
      expect(sheet2.frozenColumns, equals(2));
    });

    test('No freeze panes configuration outputs standard clean view', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.frozenRows = null;
      sheet.frozenColumns = 0; // 0 or null should behave as no freeze panes

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      expect(sheet2.frozenRows, isNull);
      expect(sheet2.frozenColumns, isNull);
    });
  });
}
