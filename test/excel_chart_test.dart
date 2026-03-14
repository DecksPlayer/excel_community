import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:archive/archive.dart';

void main() {
  test('Create Excel with ColumnChart and verify ZIP integrity', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    // Add some data
    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value"));
    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("A"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(10));

    // Add a Column Chart
    var chart = ColumnChart(
      title: "Test Chart",
      series: [
        ChartSeries(
          name: "Series 1",
          categoriesRange: r"Sheet1!$A$2:$A$2",
          valuesRange: r"Sheet1!$B$2:$B$2",
        ),
      ],
      anchor: ChartAnchor.at(column: 4, row: 1),
    );

    sheet.addChart(chart);

    var bytes = excel.save();
    expect(bytes, isNotNull);

    // Verify ZIP integrity
    var archive = ZipDecoder().decodeBytes(bytes!);
    bool foundDrawing = archive.files.any((f) => f.name == 'xl/drawings/drawing1.xml');
    bool foundChart = archive.files.any((f) => f.name == 'xl/charts/chart1.xml');
    
    expect(foundDrawing, isTrue, reason: 'drawing1.xml must be present');
    expect(foundChart, isTrue, reason: 'chart1.xml must be present');

    // Re-decode to check if it crashes
    var excel2 = Excel.decodeBytes(bytes);
    expect(excel2.sheets.containsKey('Sheet1'), isTrue);
  });

  test('Verify multi-sheet drawing integrity', () {
    var excel = Excel.createExcel();
    
    // Sheet 1 with chart
    var sheet1 = excel['Sheet1'];
    sheet1.updateCell(CellIndex.indexByString("A1"), IntCellValue(10));
    sheet1.addChart(ColumnChart(
      title: "C1",
      series: [ChartSeries(name: "S1", categoriesRange: r'Sheet1!$A$1:$A$1', valuesRange: r'Sheet1!$A$1:$A$1')],
      anchor: ChartAnchor.at(column: 2, row: 2),
    ));

    // Sheet 2 with chart (Ensure Sheet 1 is preserved and not auto-renamed)
    var sheet2 = excel['Sheet2'];
    sheet2.updateCell(CellIndex.indexByString("A1"), IntCellValue(20));
    sheet2.addChart(ColumnChart(
      title: "C2",
      series: [ChartSeries(name: "S2", categoriesRange: r'Sheet2!$A$1:$A$1', valuesRange: r'Sheet2!$A$1:$A$1')],
      anchor: ChartAnchor.at(column: 2, row: 2),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    
    expect(archive.files.any((f) => f.name == 'xl/drawings/drawing1.xml'), isTrue);
    expect(archive.files.any((f) => f.name == 'xl/drawings/drawing2.xml'), isTrue);
    expect(archive.files.any((f) => f.name == 'xl/charts/chart1.xml'), isTrue);
    expect(archive.files.any((f) => f.name == 'xl/charts/chart2.xml'), isTrue);
  });
}
