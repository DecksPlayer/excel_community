// Cell comments snippet
library;

// ---------------------------------------------------------------------------
// Cell comments
// ---------------------------------------------------------------------------
const String cellCommentsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateExcelWithComments() {
  // Create a new Excel workbook
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Add some data and comments
  var cellA1 = sheet.cell(CellIndex.indexByString("A1"));
  cellA1.value = TextCellValue("Product A");
  // Set comment on cell A1
  cellA1.comment = "This product has a 10% discount this month.";

  var cellB1 = sheet.cell(CellIndex.indexByString("B1"));
  cellB1.value = IntCellValue(150);
  // Set comment on cell B1
  cellB1.comment = "Target sales exceeded by 20 units.";

  var cellA2 = sheet.cell(CellIndex.indexByString("A2"));
  cellA2.value = TextCellValue("Product B");
  cellA2.comment = "Discontinued next quarter.";

  var cellB2 = sheet.cell(CellIndex.indexByString("B2"));
  cellB2.value = IntCellValue(85);

  // You can also read a cell's comment
  String? comment = sheet.cell(CellIndex.indexByString("A1")).comment;
  print("Cell A1 comment: \$comment");

  // Save the workbook
  excel.save(fileName: 'cell_comments_demo.xlsx');
}
''';
