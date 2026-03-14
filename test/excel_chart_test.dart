import 'package:excel_community/excel.dart';
import 'package:test/test.dart';

void main() {
  test('Create Excel with ColumnChart', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    // Add some data
    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value"));
    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("A"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(10));
    sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("B"));
    sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(20));
    sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("C"));
    sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(30));

    // Add a Column Chart
    var chart = ColumnChart(
      title: "Test Chart",
      series: [
        ChartSeries(
          name: "Series 1",
          categoriesRange: r"Sheet1!$A$2:$A$4",
          valuesRange: r"Sheet1!$B$2:$B$4",
        ),
      ],
      anchor: ChartAnchor.at(column: 4, row: 1),
    );

    sheet.addChart(chart);

    var bytes = excel.save();
    expect(bytes, isNotNull);
    expect(bytes!.isNotEmpty, isTrue);

    // Re-decode to check if it crashes
    var excel2 = Excel.decodeBytes(bytes);
    expect(excel2.sheets.containsKey('Sheet1'), isTrue);
  });
}
