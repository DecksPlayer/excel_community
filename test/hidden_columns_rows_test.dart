import 'package:test/test.dart';
import 'package:excel_community/excel_community.dart';

void main() {
  group('Hidden Columns and Rows Tests', () {
    test('Set, check, and save hidden columns and rows', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      // Add some sample data
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Visible Col A"));
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Hidden Col B"));
      sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Visible Col C"));
      
      sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Visible Row 2"));
      sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Hidden Row 3"));
      sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("Visible Row 4"));

      // Set hidden column (Column B, index 1)
      sheet.setColumnHidden(1, true);
      expect(sheet.isColumnHidden(1), isTrue);
      expect(sheet.isColumnHidden(0), isFalse);

      // Set hidden row (Row 3, index 2)
      sheet.setRowHidden(2, true);
      expect(sheet.isRowHidden(2), isTrue);
      expect(sheet.isRowHidden(1), isFalse);

      // Save and decode
      final bytes = excel.encode()!;
      final decoded = Excel.decodeBytes(bytes);
      final decodedSheet = decoded['Sheet1'];

      // Verify states are correctly parsed/preserved
      expect(decodedSheet.isColumnHidden(1), isTrue);
      expect(decodedSheet.isColumnHidden(0), isFalse);
      expect(decodedSheet.isColumnHidden(2), isFalse);

      expect(decodedSheet.isRowHidden(2), isTrue);
      expect(decodedSheet.isRowHidden(1), isFalse);
      expect(decodedSheet.isRowHidden(3), isFalse);
    });

    test('Toggle hidden state back to visible', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.setColumnHidden(2, true);
      expect(sheet.isColumnHidden(2), isTrue);

      sheet.setColumnHidden(2, false);
      expect(sheet.isColumnHidden(2), isFalse);

      sheet.setRowHidden(5, true);
      expect(sheet.isRowHidden(5), isTrue);

      sheet.setRowHidden(5, false);
      expect(sheet.isRowHidden(5), isFalse);
    });
  });
}
