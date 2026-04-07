import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:xml/xml.dart';

void main() {
  print('=== DIAGNÓSTICO DE UNDERLINE ===\n');

  // Crear Excel con underline
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Single'),
    cellStyle: CellStyle(underline: Underline.Single, fontSize: 14),
  );

  sheet.updateCell(
    CellIndex.indexByString('B1'),
    TextCellValue('Double'),
    cellStyle: CellStyle(underline: Underline.Double, fontSize: 14),
  );

  sheet.updateCell(
    CellIndex.indexByString('C1'),
    TextCellValue('None'),
    cellStyle: CellStyle(underline: Underline.None, fontSize: 14),
  );

  // Guardar
  var bytes = excel.encode()!;
  File('debug_underline.xlsx').writeAsBytesSync(bytes);

  print('Archivo guardado: debug_underline.xlsx');
  print('Tamaño: ${bytes.length} bytes\n');

  // Extraer y leer el XML de estilos
  final archive = ZipDecoder().decodeBytes(bytes);
  final stylesFile = archive.findFile('xl/styles.xml');

  if (stylesFile != null) {
    stylesFile.decompress();
    var xmlString = String.fromCharCodes(stylesFile.content);
    var document = XmlDocument.parse(xmlString);

    print('=== FONTS EN EL XML ===\n');
    var fonts = document.findAllElements('font');
    var fontIndex = 0;
    for (var font in fonts) {
      print('Font $fontIndex:');
      print('  Bold: ${font.findElements('b').isNotEmpty}');
      print('  Italic: ${font.findElements('i').isNotEmpty}');

      var underlineElements = font.findElements('u');
      if (underlineElements.isNotEmpty) {
        var u = underlineElements.first;
        var val = u.getAttribute('val');
        if (val != null) {
          print('  Underline: Double (val="$val")');
        } else {
          print('  Underline: Single (sin atributo val)');
        }
      } else {
        print('  Underline: None');
      }

      var size = font.findElements('sz');
      if (size.isNotEmpty) {
        print('  Size: ${size.first.getAttribute('val')}');
      }
      print('');
      fontIndex++;
    }

    print('=== CELL XF (estilos de celda) ===\n');
    var cellXfs = document.findAllElements('cellXfs').first;
    var xfs = cellXfs.findElements('xf');
    var xfIndex = 0;
    for (var xf in xfs) {
      var fontId = xf.getAttribute('fontId') ?? '0';
      var numFmtId = xf.getAttribute('numFmtId') ?? '0';
      print('XF $xfIndex: fontId=$fontId, numFmtId=$numFmtId');
      xfIndex++;
    }
  }

  // Ahora leer de vuelta
  print('\n=== LEYENDO DE VUELTA ===\n');
  var excel2 = Excel.decodeBytes(bytes);

  var cellA1 = excel2['Sheet1'].cell(CellIndex.indexByString('A1'));
  var cellB1 = excel2['Sheet1'].cell(CellIndex.indexByString('B1'));
  var cellC1 = excel2['Sheet1'].cell(CellIndex.indexByString('C1'));

  print('A1 (Single esperado): ${cellA1.cellStyle?.underline}');
  print('B1 (Double esperado): ${cellB1.cellStyle?.underline}');
  print('C1 (None esperado): ${cellC1.cellStyle?.underline}');

  print('\n✅ Diagnóstico completado');
}
