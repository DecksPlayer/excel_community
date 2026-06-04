import 'package:excel_community/excel_community.dart';
import 'package:archive/archive.dart';

void main() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Cat"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Val"));
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alpha"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(10));
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Beta"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(20));

  sheet.addChart(ColumnChart(
    title: "Test Chart",
    series: [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!$A$2:$A$3",
        valuesRange: r"Sheet1!$B$2:$B$3",
      )
    ],
    anchor: ChartAnchor.at(column: 4, row: 1),
  ));

  var bytes = excel.save()!;
  var archive = ZipDecoder().decodeBytes(bytes);
  for (var f in archive.files) {
    if (f.name == 'xl/charts/chart1.xml' ||
        f.name == 'xl/drawings/drawing1.xml' ||
        f.name == 'xl/charts/_rels/chart1.xml.rels' ||
        f.name == '[Content_Types].xml') {
      print("=== ${f.name} ===");
      print(String.fromCharCodes(f.content));
      print('');
    }
  }
}
