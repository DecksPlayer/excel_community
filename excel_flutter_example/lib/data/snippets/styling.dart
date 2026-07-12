// Styling snippets: number formats (currency, custom patterns).
library;

// ---------------------------------------------------------------------------
// Number formats
// ---------------------------------------------------------------------------
const String numberFormatsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateNumberFormats() {
  var excel = Excel.createExcel();
  var sheet = excel['Number Formatting'];

  // Apply numberFormat to CellStyle
  var currencyStyle = CellStyle(
    numberFormat: NumFormat.standard_44, // Currency Accounting Indent fmt
    horizontalAlign: HorizontalAlign.Right,
  );

  var cell = sheet.cell(CellIndex.indexByString("C3"));
  cell.value = DoubleCellValue(1234.56);
  cell.cellStyle = currencyStyle;

  // Custom Formats can also be declared:
  var customStyle = CellStyle(
    numberFormat: CustomNumericNumFormat('#,##0.00 \\m\\²'),
  );

  excel.save(fileName: 'number_formats_demo.xlsx');
}
''';
