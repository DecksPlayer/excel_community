import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('Sheet Protection and Cell Locking Tests', () {
    test('Set, save, and decode protected sheet with password', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Protect sheet with password
      sheet.protect('omixomPassword');

      expect(sheet.isProtected, isTrue);
      expect(sheet.sheetProtection.sheet, isTrue);
      expect(sheet.sheetProtection.password, isNotNull);
      // 'omixomPassword' legacy 16-bit hash check:
      // We will make sure it is generated and is not empty.
      expect(sheet.sheetProtection.password!.length, equals(4));

      // Save to bytes
      var bytes = excel.encode();
      expect(bytes, isNotNull);
      expect(bytes!.isNotEmpty, isTrue);

      // Decode the bytes back to a new Excel workbook
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      // Assert properties are correctly preserved and parsed back
      expect(sheet2.isProtected, isTrue);
      expect(sheet2.sheetProtection.sheet, isTrue);
      expect(sheet2.sheetProtection.password, equals(sheet.sheetProtection.password));
    });

    test('Custom SheetProtection properties preservation', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.sheetProtection.sheet = true;
      sheet.sheetProtection.formatCells = false;
      sheet.sheetProtection.selectLockedCells = false;
      sheet.sheetProtection.selectUnlockedCells = true;

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      expect(sheet2.isProtected, isTrue);
      expect(sheet2.sheetProtection.formatCells, isFalse);
      expect(sheet2.sheetProtection.selectLockedCells, isFalse);
      expect(sheet2.sheetProtection.selectUnlockedCells, isTrue);
    });

    test('Cell Locking style preservation', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      final lockedStyle = CellStyle(locked: true, bold: true);
      final unlockedStyle = CellStyle(locked: false);

      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Locked Header"), cellStyle: lockedStyle);
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Unlocked Input"), cellStyle: unlockedStyle);

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      final cellA1 = sheet2.cell(CellIndex.indexByString("A1"));
      final cellB1 = sheet2.cell(CellIndex.indexByString("B1"));

      expect(cellA1.cellStyle?.locked, isTrue);
      expect(cellB1.cellStyle?.locked, isFalse);
    });

    test('No sheet protection elements when disabled', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.sheetProtection.sheet = false;

      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var sheet2 = excel2['Sheet1'];

      expect(sheet2.isProtected, isFalse);
    });
  });
}
