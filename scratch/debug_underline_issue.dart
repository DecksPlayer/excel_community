import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:xml/xml.dart';

void main() {
  try {
    print('Generating spreadsheet...');
    var excel = Excel.createExcel();
    var sheet = excel['Text Styles Demo'];
    excel.delete('Sheet1');

    // Título principal
    sheet.updateCell(
      CellIndex.indexByString('A1'),
      TextCellValue('EJEMPLOS DE ESTILOS DE TEXTO'),
      cellStyle: CellStyle(
        bold: true,
        fontSize: 16,
        fontColorHex: ExcelColor.blue,
        horizontalAlign: HorizontalAlign.Center,
      ),
    );
    sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('D1'));

    // === EJEMPLOS DE UNDERLINE ===
    int row = 3;

    // Header
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('UNDERLINE (SUBRAYADO)'),
      cellStyle: CellStyle(
        bold: true,
        fontSize: 14,
        backgroundColorHex: ExcelColor.grey200,
      ),
    );
    sheet.merge(
      CellIndex.indexByString('A$row'),
      CellIndex.indexByString('D$row'),
    );
    row += 2;

    // Sin subrayado
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Sin subrayado:'),
      cellStyle: CellStyle(bold: true),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue('Este texto NO tiene subrayado'),
      cellStyle: CellStyle(underline: Underline.None, fontSize: 12),
    );
    row++;

    // Subrayado simple
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Subrayado Simple:'),
      cellStyle: CellStyle(bold: true),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue('Este texto tiene subrayado SIMPLE'),
      cellStyle: CellStyle(
        underline: Underline.Single,
        fontSize: 12,
        fontColorHex: ExcelColor.blue,
      ),
    );
    row++;

    // Subrayado doble
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Subrayado Doble:'),
      cellStyle: CellStyle(bold: true),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue('Este texto tiene subrayado DOBLE'),
      cellStyle: CellStyle(
        underline: Underline.Double,
        fontSize: 12,
        fontColorHex: ExcelColor.red,
      ),
    );
    row += 2;

    // === COMBINACIONES ===
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('COMBINACIONES DE ESTILOS'),
      cellStyle: CellStyle(
        bold: true,
        fontSize: 14,
        backgroundColorHex: ExcelColor.grey200,
      ),
    );
    sheet.merge(
      CellIndex.indexByString('A$row'),
      CellIndex.indexByString('D$row'),
    );
    row += 2;

    // Bold + Underline
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Negrita + Subrayado Simple'),
      cellStyle: CellStyle(
        bold: true,
        underline: Underline.Single,
        fontSize: 12,
      ),
    );
    row++;

    // Italic + Underline
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Cursiva + Subrayado Doble'),
      cellStyle: CellStyle(
        italic: true,
        underline: Underline.Double,
        fontSize: 12,
      ),
    );
    row++;

    // Bold + Italic + Underline
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Negrita + Cursiva + Subrayado'),
      cellStyle: CellStyle(
        bold: true,
        italic: true,
        underline: Underline.Single,
        fontSize: 12,
        fontColorHex: ExcelColor.purple,
      ),
    );
    row++;

    // Strikethrough
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Texto Tachado (Strikethrough)'),
      cellStyle: CellStyle(
        strikethrough: true,
        fontSize: 12,
        fontColorHex: ExcelColor.red,
      ),
    );
    row++;

    // Strikethrough + Underline
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Tachado + Subrayado'),
      cellStyle: CellStyle(
        strikethrough: true,
        underline: Underline.Single,
        fontSize: 12,
      ),
    );
    row++;

    // All styles combined
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Negrita + Cursiva + Tachado + Subrayado'),
      cellStyle: CellStyle(
        bold: true,
        italic: true,
        strikethrough: true,
        underline: Underline.Double,
        fontSize: 12,
        fontColorHex: ExcelColor.green,
      ),
    );
    row++;

    // Color de fondo + Underline
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Fondo Amarillo + Subrayado'),
      cellStyle: CellStyle(
        underline: Underline.Single,
        fontSize: 12,
        backgroundColorHex: ExcelColor.yellow,
        fontColorHex: ExcelColor.black,
      ),
    );
    row += 2;

    // === TABLA CON BORDES ===
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('TABLA CON BORDES Y ESTILOS'),
      cellStyle: CellStyle(
        bold: true,
        fontSize: 14,
        backgroundColorHex: ExcelColor.grey200,
      ),
    );
    sheet.merge(
      CellIndex.indexByString('A$row'),
      CellIndex.indexByString('D$row'),
    );
    row += 2;

    // Encabezados de tabla
    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.blue300,
      fontColorHex: ExcelColor.white,
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue('Producto'),
      cellStyle: headerStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue('Cantidad'),
      cellStyle: headerStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue('Precio'),
      cellStyle: headerStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('D$row'),
      TextCellValue('Total'),
      cellStyle: headerStyle,
    );
    row++;

    // Filas de datos con diferentes estilos
    final cellBorderStyle = CellStyle(
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
    );

    final data = [
      ['Laptop', 2, 1200.50, 2401.00],
      ['Mouse', 5, 25.99, 129.95],
      ['Teclado', 3, 89.90, 269.70],
    ];

    for (var rowData in data) {
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue(rowData[0] as String),
        cellStyle: cellBorderStyle.copyWith(underlineVal: Underline.Single),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        IntCellValue(rowData[1] as int),
        cellStyle: cellBorderStyle.copyWith(
          horizontalAlignVal: HorizontalAlign.Center,
        ),
      );
      sheet.updateCell(
        CellIndex.indexByString('C$row'),
        DoubleCellValue(rowData[2] as double),
        cellStyle: cellBorderStyle.copyWith(
          horizontalAlignVal: HorizontalAlign.Right,
          numberFormat: NumFormat.standard_2,
        ),
      );
      sheet.updateCell(
        CellIndex.indexByString('D$row'),
        DoubleCellValue(rowData[3] as double),
        cellStyle: cellBorderStyle.copyWith(
          horizontalAlignVal: HorizontalAlign.Right,
          numberFormat: NumFormat.standard_2,
          boldVal: true,
          fontColorHexVal: ExcelColor.green,
        ),
      );
      row++;
    }

    // Ajustar anchos de columna
    sheet.setColumnWidth(0, 25.0);
    sheet.setColumnWidth(1, 15.0);
    sheet.setColumnWidth(2, 15.0);
    sheet.setColumnWidth(3, 15.0);

    final bytes = excel.encode()!;
    print('Encoded file size: ${bytes.length} bytes');

    // Decode zip and validate XML files
    print('Validating XML files inside the zip...');
    final archive = ZipDecoder().decodeBytes(bytes);
    for (final file in archive.files) {
      if (file.isFile && file.name.endsWith('.xml')) {
        final content = utf8.decode(file.content);
        try {
          XmlDocument.parse(content);
          print('✅ ${file.name} is well-formed XML');
          
          if (file.name == 'xl/styles.xml') {
            print('--- xl/styles.xml content ---');
            print(content);
            print('----------------------------');
          }
          if (file.name == 'xl/sharedStrings.xml') {
            print('--- xl/sharedStrings.xml content ---');
            print(content);
            print('----------------------------');
          }
          if (file.name == 'xl/worksheets/sheet1.xml') {
            print('--- xl/worksheets/sheet1.xml content ---');
            print(content);
            print('----------------------------');
          }
        } catch (e) {
          print('❌ ${file.name} has XML parsing errors: $e');
          print('Content snippet around error:');
          print(content.substring(0, content.length > 500 ? 500 : content.length));
        }
      }
    }
  } catch (e, st) {
    print('Failed: $e');
    print(st);
  }
}
