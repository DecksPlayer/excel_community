import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:archive/archive.dart';

void main() {
  test('Verify all files are included in the ZIP archive', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];
    
    // Add a chart which triggers drawing generation
    Chart chart = ColumnChart(
      title: "Test Chart",
      series: [
        ChartSeries(
          name: "S1",
          categoriesRange: r"Sheet1!$A$2:$A$7",
          valuesRange: r"Sheet1!$B$2:$B$7",
        ),
      ],
      anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
    );
    sheet.addChart(chart);
    
    // Add another sheet with a chart
    var sheet2 = excel['Sheet2'];
    sheet2.updateCell(CellIndex.indexByString("A1"), TextCellValue("Data"));
    sheet2.updateCell(CellIndex.indexByString("A2"), IntCellValue(100));
    
    Chart chart2 = ColumnChart(
      title: "Test Chart 2",
      series: [
        ChartSeries(
          name: "S2",
          categoriesRange: r"Sheet2!$A$1:$A$1",
          valuesRange: r"Sheet2!$A$2:$A$2",
        ),
      ],
      anchor: ChartAnchor.at(column: 2, row: 1, width: 5, height: 5),
    );
    sheet2.addChart(chart2);

    var bytes = excel.encode();
    expect(bytes, isNotNull);

    var archive = ZipDecoder().decodeBytes(bytes!);
    
    // Check for drawing and chart files
    bool foundDrawing1 = false;
    bool foundDrawing2 = false;
    bool foundChart1 = false;
    bool foundChart2 = false;
    
    for (var file in archive.files) {
      if (file.name == 'xl/drawings/drawing1.xml') foundDrawing1 = true;
      if (file.name == 'xl/drawings/drawing2.xml') foundDrawing2 = true;
      if (file.name == 'xl/charts/chart1.xml') foundChart1 = true;
      if (file.name == 'xl/charts/chart2.xml') foundChart2 = true;
    }
    
    expect(foundDrawing1, isTrue, reason: 'drawing1.xml should be in the archive');
    expect(foundDrawing2, isTrue, reason: 'drawing2.xml should be in the archive');
    expect(foundChart1, isTrue, reason: 'chart1.xml should be in the archive');
    expect(foundChart2, isTrue, reason: 'chart2.xml should be in the archive');
  });
}
