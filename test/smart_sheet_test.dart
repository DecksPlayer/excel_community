import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  test('Smart Sheet Initialization: accessing new sheet renames Sheet1', () {
    var excel = Excel.createExcel();

    // Check initial state
    expect(excel.sheets.keys.length, equals(1));
    expect(excel.sheets.containsKey('Sheet1'), isTrue);

    // Access a new sheet
    var mySheet = excel['MySheet'];

    // Should have renamed Sheet1 to MySheet
    expect(excel.sheets.keys.length, equals(1));
    expect(excel.sheets.containsKey('MySheet'), isTrue);
    expect(excel.sheets.containsKey('Sheet1'), isFalse);

    mySheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Hello"));

    // Now access another one, it should ADD this time because MySheet is not empty or not Sheet1
    excel['AnotherSheet'];
    expect(excel.sheets.keys.length, equals(2));
    expect(excel.sheets.containsKey('MySheet'), isTrue);
    expect(excel.sheets.containsKey('AnotherSheet'), isTrue);
  });

  test('Sheet1 requested explicitly is not renamed by a later new sheet', () {
    final excel = Excel.createExcel();
    final sheet1 = excel['Sheet1'];
    excel['Other'];
    sheet1.updateCell(CellIndex.indexByString('A1'), TextCellValue('kept'));

    expect(excel.sheets.keys, ['Sheet1', 'Other']);
    final decoded = Excel.decodeBytes(excel.encode()!);
    expect(decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value, TextCellValue('kept'));
  });
}
