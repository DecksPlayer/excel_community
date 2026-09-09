import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('Data.displayText', () {
    test('empty cell renders as empty string', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));

      expect(cell.displayText, equals(''));
    });

    test('General format renders ints without style noise', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final intCell = sheet.cell(CellIndex.indexByString('A1'));
      intCell.value = IntCellValue(42);
      expect(intCell.displayText, equals('42'));
    });

    test('General format renders a double with the General format code', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final doubleCell = sheet.cell(CellIndex.indexByString('A1'));
      doubleCell.value = DoubleCellValue(12.3);
      doubleCell.cellStyle = CellStyle(numberFormat: NumFormat.standard_0);
      expect(doubleCell.displayText, equals('12.3'));
    });

    test('thousands separator format (#,##0)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = IntCellValue(1234567);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_3);

      expect(cell.displayText, equals('1,234,567'));
    });

    test('fixed 2-decimal format (0.00)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(3.5);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_2);

      expect(cell.displayText, equals('3.50'));
    });

    test(r'currency format ($#,##0.00_);($#,##0.00))', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final positive = sheet.cell(CellIndex.indexByString('A1'));
      positive.value = DoubleCellValue(1234.5);
      positive.cellStyle = CellStyle(numberFormat: NumFormat.standard_7);
      expect(positive.displayText, equals(r'$1,234.50'));

      final negative = sheet.cell(CellIndex.indexByString('A2'));
      negative.value = DoubleCellValue(-1234.5);
      negative.cellStyle = CellStyle(numberFormat: NumFormat.standard_7);
      expect(negative.displayText, equals(r'($1,234.50)'));
    });

    test('percentage format (0.00%)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(0.4567);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_10);

      expect(cell.displayText, equals('45.67%'));
    });

    test('date format (mm-dd-yy)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DateCellValue(year: 2023, month: 4, day: 20);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_14);

      expect(cell.displayText, equals('04-20-23'));
    });

    test('date/time format (m/d/yy h:mm)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DateTimeCellValue(
          year: 2023, month: 4, day: 20, hour: 15, minute: 44, second: 13);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_22);

      expect(cell.displayText, equals('4/20/23 15:44'));
    });

    test('time format with AM/PM (h:mm AM/PM)', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = TimeCellValue(hour: 14, minute: 5, second: 0);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_18);

      expect(cell.displayText, equals('2:05 PM'));
    });

    test('elapsed-time format ([h]:mm:ss) preserves hours over 24', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = TimeCellValue(hour: 30, minute: 6, second: 5);
      cell.cellStyle =
          CellStyle(numberFormat: NumFormat.standard_46); // [h]:mm:ss

      expect(cell.displayText, equals('30:06:05'));
    });

    test('formula cell renders its cached value, not the formula text', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value =
          const FormulaCellValue('SUM(1,2)', cachedValue: IntCellValue(3));
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_2);

      expect(cell.displayText, equals('3.00'));
    });

    test('formula cell with no cached value renders empty', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = const FormulaCellValue('SUM(1,2)');

      expect(cell.displayText, equals(''));
    });

    test('custom currency-like format via NumFormat.custom', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(999.9);
      cell.cellStyle =
          CellStyle(numberFormat: NumFormat.custom(formatCode: '#,##0.0"kg"'));

      expect(cell.displayText, equals('999.9kg'));
    });
  });
}
