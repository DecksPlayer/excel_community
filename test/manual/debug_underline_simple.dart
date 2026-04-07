import 'package:excel_community/excel_community.dart';

void main() {
  print('=== DIAGNÓSTICO PROFUNDO DE UNDERLINE ===\n');

  // Crear Excel con underline
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Crear CellStyle con underline
  var styleWithUnderline = CellStyle(
    underline: Underline.Single,
    fontSize: 14,
  );

  print('CellStyle creado:');
  print('  Underline: ${styleWithUnderline.underline}');
  print('  Font Size: ${styleWithUnderline.fontSize}');
  print('');

  // Aplicar a celda
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Test'),
    cellStyle: styleWithUnderline,
  );

  // Verificar que la celda tiene el estilo
  var cell = sheet.cell(CellIndex.indexByString('A1'));
  print('Celda A1 después de updateCell:');
  print('  Underline: ${cell.cellStyle?.underline}');
  print('  Font Size: ${cell.cellStyle?.fontSize}');
  print('');

  // Intentar guardar y ver qué pasa
  print('Guardando...');
  var bytes = excel.encode();

  if (bytes == null) {
    print('ERROR: encode() devolvió null');
  } else {
    print('Bytes: ${bytes.length}');
  }
}
