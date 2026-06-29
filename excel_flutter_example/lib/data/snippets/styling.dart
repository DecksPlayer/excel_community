// Styling snippets: text styles (bold, italic, underline, colors) and
// number formats (currency, custom patterns).
library;

// ---------------------------------------------------------------------------
// Text styles
// ---------------------------------------------------------------------------
const String textStylesSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateTextStyles() {
  var excel = Excel.createExcel();
  var sheet = excel['Styles Demo'];

  // Define CellStyle configurations
  var cellStyle = CellStyle(
    bold: true,
    italic: true,
    underline: Underline.Double, // Double underline
    strikethrough: true,        // Strikethrough line
    fontSize: 12,
    fontColorHex: ExcelColor.fromHexString('#FF0000'),       // Red text
    backgroundColorHex: ExcelColor.fromHexString('#CCCCCC'), // Grey fill
  );

  var cell = sheet.cell(CellIndex.indexByString("B5"));
  cell.value = TextCellValue("Styled Text Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'text_styles_demo.xlsx');
}
''';

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