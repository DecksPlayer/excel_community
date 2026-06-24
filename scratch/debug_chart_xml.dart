import 'dart:convert';
import 'package:excel_community/excel_community.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Add sample data
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value"));
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alpha"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(10));
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Beta"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(20));

  // Add a Column Chart
  var chart = ColumnChart(
    title: "Test Chart",
    series: [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!$A$2:$A$3",
        valuesRange: r"Sheet1!$B$2:$B$3",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1),
  );

  sheet.addChart(chart);

  var bytes = excel.save();
  if (bytes == null) {
    print("Error: Excel save returned null");
    return;
  }

  var archive = ZipDecoder().decodeBytes(bytes);

  print("=== ZIP FILES ===");
  for (var file in archive.files) {
    if (file.name.contains('drawing') ||
        file.name.contains('chart') ||
        file.name.contains('sheet1') ||
        file.name.contains('[Content_Types]')) {
      print("- ${file.name} (${file.content.length} bytes)");
    }
  }

  void printXmlFile(String name) {
    var file = archive.files.firstWhere((f) => f.name == name,
        orElse: () => throw "File $name not found");
    var xmlStr = utf8.decode(file.content);
    print("\n=== XML: $name ===");
    try {
      var document = XmlDocument.parse(xmlStr);
      print(document.toXmlString(pretty: true));
    } catch (e) {
      print("PARSE ERROR: $e");
      print(xmlStr);
    }
  }

  printXmlFile('[Content_Types].xml');
  printXmlFile('xl/worksheets/sheet1.xml');
  printXmlFile('xl/worksheets/_rels/sheet1.xml.rels');
  printXmlFile('xl/drawings/drawing1.xml');
  printXmlFile('xl/drawings/_rels/drawing1.xml.rels');
  printXmlFile('xl/charts/chart1.xml');
  printXmlFile('xl/charts/_rels/chart1.xml.rels');
}
