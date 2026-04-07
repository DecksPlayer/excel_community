import 'dart:io';
import 'package:excel_community/excel_community.dart';

void main() {
  print('=== GENERANDO EXCEL DE EJEMPLO CON UNDERLINE ===\n');

  // Crear un nuevo Excel
  var excel = Excel.createExcel();
  var sheet = excel['Ejemplos de Estilos'];

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

  // Merge cells para el título
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
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('D$row'));
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
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('D$row'));
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
    TextCellValue('Negrita + Cursiva + Subrayado Simple'),
    cellStyle: CellStyle(
      bold: true,
      italic: true,
      underline: Underline.Single,
      fontSize: 12,
      fontColorHex: ExcelColor.purple,
    ),
  );
  row++;

  // Color de fondo + Underline
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Fondo Amarillo + Subrayado Simple'),
    cellStyle: CellStyle(
      underline: Underline.Single,
      fontSize: 12,
      backgroundColorHex: ExcelColor.yellow,
      fontColorHex: ExcelColor.black,
    ),
  );
  row += 2;

  // === OTROS ESTILOS ===
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('OTROS ESTILOS'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('D$row'));
  row += 2;

  // Diferentes tamaños
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Texto tamaño 10'),
    cellStyle: CellStyle(fontSize: 10),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Texto tamaño 14'),
    cellStyle: CellStyle(fontSize: 14),
  );
  sheet.updateCell(
    CellIndex.indexByString('C$row'),
    TextCellValue('Texto tamaño 18'),
    cellStyle: CellStyle(fontSize: 18),
  );
  row++;

  // Diferentes colores
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Texto Rojo'),
    cellStyle: CellStyle(fontColorHex: ExcelColor.red, fontSize: 12),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Texto Verde'),
    cellStyle: CellStyle(fontColorHex: ExcelColor.green, fontSize: 12),
  );
  sheet.updateCell(
    CellIndex.indexByString('C$row'),
    TextCellValue('Texto Azul'),
    cellStyle: CellStyle(fontColorHex: ExcelColor.blue, fontSize: 12),
  );
  row++;

  // Alineaciones
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Izquierda'),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      fontSize: 12,
    ),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Centro'),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      fontSize: 12,
    ),
  );
  sheet.updateCell(
    CellIndex.indexByString('C$row'),
    TextCellValue('Derecha'),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Right,
      fontSize: 12,
    ),
  );
  row += 2;

  // Tabla con bordes
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('TABLA CON BORDES'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('D$row'));
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

  // Filas de datos
  final cellStyle = CellStyle(
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  final data = [
    ['Laptop', '2', '1200.50', '2401.00'],
    ['Mouse', '5', '25.99', '129.95'],
    ['Teclado', '3', '89.90', '269.70'],
  ];

  for (var rowData in data) {
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue(rowData[0]),
      cellStyle: cellStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      IntCellValue(int.parse(rowData[1])),
      cellStyle: cellStyle.copyWith(horizontalAlignVal: HorizontalAlign.Center),
    );
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      DoubleCellValue(double.parse(rowData[2])),
      cellStyle: cellStyle.copyWith(
        horizontalAlignVal: HorizontalAlign.Right,
        numberFormat: NumFormat.standard_2,
      ),
    );
    sheet.updateCell(
      CellIndex.indexByString('D$row'),
      DoubleCellValue(double.parse(rowData[3])),
      cellStyle: cellStyle.copyWith(
        horizontalAlignVal: HorizontalAlign.Right,
        numberFormat: NumFormat.standard_2,
        boldVal: true,
      ),
    );
    row++;
  }

  // Ajustar anchos de columna
  sheet.setColumnWidth(0, 25.0); // A
  sheet.setColumnWidth(1, 15.0); // B
  sheet.setColumnWidth(2, 15.0); // C
  sheet.setColumnWidth(3, 15.0); // D

  // Eliminar Sheet1 por defecto
  excel.delete('Sheet1');

  // Guardar archivo
  final outputFile = File('example/ejemplo_underline_estilos.xlsx');
  final bytes = excel.encode();

  if (bytes != null) {
    outputFile
      ..createSync(recursive: true)
      ..writeAsBytesSync(bytes);

    print('✅ Archivo creado exitosamente!');
    print('📁 Ubicación: ${outputFile.path}');
    print('📊 Tamaño: ${bytes.length} bytes');
    print('');
    print('Puedes abrir este archivo con:');
    print('  - Microsoft Excel');
    print('  - Google Sheets');
    print('  - LibreOffice Calc');
    print('  - El ejemplo de Flutter');
    print('');
    print('El archivo contiene:');
    print('  ✅ Ejemplos de underline (simple y doble)');
    print('  ✅ Combinaciones de estilos');
    print('  ✅ Diferentes colores y tamaños');
    print('  ✅ Tabla con bordes');
    print('  ✅ Alineaciones');
    print('');
    print('¡Ahora puedes probarlo en la app de Flutter!');
  } else {
    print('❌ Error al generar el archivo');
  }
}
