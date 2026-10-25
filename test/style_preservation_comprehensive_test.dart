import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Style Preservation Tests', () {
    test('Complete style preservation after read/write cycle', () {
      // Crear un Excel con varios estilos
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Estilo completo con todas las propiedades
      var completeStyle = CellStyle(
        // Font properties
        fontColorHex: ExcelColor.red,
        fontFamily: 'Arial',
        fontSize: 14,
        bold: true,
        italic: true,
        underline: Underline.Single,
        // Background
        backgroundColorHex: ExcelColor.yellow,
        // Alignment
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        rotation: 45,
        textWrapping: TextWrapping.WrapText,
        // Borders
        leftBorder: Border(
          borderStyle: BorderStyle.Thin,
          borderColorHex: ExcelColor.blue,
        ),
        rightBorder: Border(
          borderStyle: BorderStyle.Medium,
          borderColorHex: ExcelColor.green,
        ),
        topBorder: Border(
          borderStyle: BorderStyle.Thick,
          borderColorHex: ExcelColor.black,
        ),
        bottomBorder: Border(
          borderStyle: BorderStyle.Dashed,
          borderColorHex: ExcelColor.grey200,
        ),
        // Number format
        numberFormat: NumFormat.standard_2,
      );

      // Aplicar el estilo a una celda
      var cellIndex = CellIndex.indexByString('B2');
      sheet.updateCell(
        cellIndex,
        DoubleCellValue(123.45),
        cellStyle: completeStyle,
      );

      // Guardar y recargar
      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);
      var cell2 = excel2['Sheet1'].cell(cellIndex);
      var style2 = cell2.cellStyle;

      // Verificar que el estilo se conserva
      expect(style2, isNotNull, reason: 'Cell style should not be null');

      print('\n=== VERIFICACIÓN DE PRESERVACIÓN DE ESTILOS ===\n');

      // Font properties
      print('Font Color: ${completeStyle.fontColor} -> ${style2?.fontColor}');
      expect(
        style2?.fontColor,
        equals(completeStyle.fontColor),
        reason: 'Font color should be preserved',
      );

      print(
          'Font Family: ${completeStyle.fontFamily} -> ${style2?.fontFamily}');
      expect(
        style2?.fontFamily,
        equals(completeStyle.fontFamily),
        reason: 'Font family should be preserved',
      );

      print('Font Size: ${completeStyle.fontSize} -> ${style2?.fontSize}');
      expect(
        style2?.fontSize,
        equals(completeStyle.fontSize),
        reason: 'Font size should be preserved',
      );

      print('Bold: ${completeStyle.isBold} -> ${style2?.isBold}');
      expect(
        style2?.isBold,
        equals(completeStyle.isBold),
        reason: 'Bold should be preserved',
      );

      print('Italic: ${completeStyle.isItalic} -> ${style2?.isItalic}');
      expect(
        style2?.isItalic,
        equals(completeStyle.isItalic),
        reason: 'Italic should be preserved',
      );

      print('Underline: ${completeStyle.underline} -> ${style2?.underline}');
      expect(
        style2?.underline,
        equals(completeStyle.underline),
        reason: 'Underline should be preserved',
      );

      // Background
      print(
          'Background: ${completeStyle.backgroundColor} -> ${style2?.backgroundColor}');
      expect(
        style2?.backgroundColor,
        equals(completeStyle.backgroundColor),
        reason: 'Background color should be preserved',
      );

      // Alignment
      print(
          'Horizontal Align: ${completeStyle.horizontalAlignment} -> ${style2?.horizontalAlignment}');
      expect(
        style2?.horizontalAlignment,
        equals(completeStyle.horizontalAlignment),
        reason: 'Horizontal alignment should be preserved',
      );

      print(
          'Vertical Align: ${completeStyle.verticalAlignment} -> ${style2?.verticalAlignment}');
      expect(
        style2?.verticalAlignment,
        equals(completeStyle.verticalAlignment),
        reason: 'Vertical alignment should be preserved',
      );

      print('Rotation: ${completeStyle.rotation} -> ${style2?.rotation}');
      expect(
        style2?.rotation,
        equals(completeStyle.rotation),
        reason: 'Rotation should be preserved',
      );

      print('Text Wrapping: ${completeStyle.wrap} -> ${style2?.wrap}');
      expect(
        style2?.wrap,
        equals(completeStyle.wrap),
        reason: 'Text wrapping should be preserved',
      );

      // Borders
      print(
          'Left Border Style: ${completeStyle.leftBorder.borderStyle} -> ${style2?.leftBorder.borderStyle}');
      expect(
        style2?.leftBorder.borderStyle,
        equals(completeStyle.leftBorder.borderStyle),
        reason: 'Left border style should be preserved',
      );

      print(
          'Left Border Color: ${completeStyle.leftBorder.borderColorHex} -> ${style2?.leftBorder.borderColorHex}');
      expect(
        style2?.leftBorder.borderColorHex,
        equals(completeStyle.leftBorder.borderColorHex),
        reason: 'Left border color should be preserved',
      );

      print(
          'Right Border Style: ${completeStyle.rightBorder.borderStyle} -> ${style2?.rightBorder.borderStyle}');
      expect(
        style2?.rightBorder.borderStyle,
        equals(completeStyle.rightBorder.borderStyle),
        reason: 'Right border style should be preserved',
      );

      print(
          'Top Border Style: ${completeStyle.topBorder.borderStyle} -> ${style2?.topBorder.borderStyle}');
      expect(
        style2?.topBorder.borderStyle,
        equals(completeStyle.topBorder.borderStyle),
        reason: 'Top border style should be preserved',
      );

      print(
          'Bottom Border Style: ${completeStyle.bottomBorder.borderStyle} -> ${style2?.bottomBorder.borderStyle}');
      expect(
        style2?.bottomBorder.borderStyle,
        equals(completeStyle.bottomBorder.borderStyle),
        reason: 'Bottom border style should be preserved',
      );

      // Number format
      print(
          'Number Format: ${completeStyle.numberFormat} -> ${style2?.numberFormat}');
      expect(
        style2?.numberFormat,
        equals(completeStyle.numberFormat),
        reason: 'Number format should be preserved',
      );

      print('\n✅ TODOS LOS ESTILOS SE PRESERVAN CORRECTAMENTE\n');
    });

    test('Multiple cells with different styles', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Crear múltiples celdas con estilos diferentes
      var styles = <String, CellStyle>{
        'A1': CellStyle(
          bold: true,
          fontSize: 16,
          horizontalAlign: HorizontalAlign.Left,
        ),
        'B1': CellStyle(
          italic: true,
          fontSize: 12,
          horizontalAlign: HorizontalAlign.Center,
        ),
        'C1': CellStyle(
          underline: Underline.Double,
          fontSize: 10,
          horizontalAlign: HorizontalAlign.Right,
        ),
        'A2': CellStyle(
          fontColorHex: ExcelColor.red,
          backgroundColorHex: ExcelColor.yellow,
        ),
        'B2': CellStyle(
          fontColorHex: ExcelColor.blue,
          backgroundColorHex: ExcelColor.green400,
        ),
      };

      // Aplicar estilos
      styles.forEach((cellRef, style) {
        sheet.updateCell(
          CellIndex.indexByString(cellRef),
          TextCellValue('Cell $cellRef'),
          cellStyle: style,
        );
      });

      // Guardar y recargar
      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);

      print('\n=== VERIFICACIÓN DE MÚLTIPLES CELDAS ===\n');

      // Verificar cada celda
      styles.forEach((cellRef, originalStyle) {
        var cell = excel2['Sheet1'].cell(CellIndex.indexByString(cellRef));
        var loadedStyle = cell.cellStyle;

        print('Celda $cellRef:');
        print('  Bold: ${originalStyle.isBold} -> ${loadedStyle?.isBold}');
        print(
            '  Italic: ${originalStyle.isItalic} -> ${loadedStyle?.isItalic}');
        print(
            '  Font Size: ${originalStyle.fontSize} -> ${loadedStyle?.fontSize}');
        print(
            '  H-Align: ${originalStyle.horizontalAlignment} -> ${loadedStyle?.horizontalAlignment}');

        expect(loadedStyle, isNotNull);
        expect(loadedStyle?.isBold, equals(originalStyle.isBold));
        expect(loadedStyle?.isItalic, equals(originalStyle.isItalic));
        // Si fontSize no se setea explícitamente, Excel usa el default (12),
        // así que comparamos solo si fue seteado
        if (originalStyle.fontSize != null) {
          expect(loadedStyle?.fontSize, equals(originalStyle.fontSize));
        }
        expect(loadedStyle?.horizontalAlignment,
            equals(originalStyle.horizontalAlignment));
      });

      print('\n✅ TODOS LOS ESTILOS DE MÚLTIPLES CELDAS SE PRESERVAN\n');
    });
  });

  test('semi-transparent background colors keep every fill in place', () {
    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];
    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('shade'),
        cellStyle: CellStyle(backgroundColorHex: ExcelColor.black12));
    sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('red'),
        cellStyle: CellStyle(backgroundColorHex: ExcelColor.red));
    final bytes = excel.encode()!;

    final styles = XmlDocument.parse(utf8.decode(
        ZipDecoder().decodeBytes(bytes).findFile('xl/styles.xml')!.content));
    final fillsElement = styles.rootElement.getElement('fills')!;
    final fills = fillsElement.findElements('fill').toList();
    expect(fillsElement.getAttribute('count'), '${fills.length}');

    final xfs = styles.rootElement.getElement('cellXfs')!.findElements('xf').toList();
    final sheetXml = utf8.decode(
        ZipDecoder().decodeBytes(bytes).findFile('xl/worksheets/sheet1.xml')!.content);
    String fillOf(String cell) {
      final s = int.parse(RegExp('<c r="$cell" s="(\\d+)"').firstMatch(sheetXml)!.group(1)!);
      final fillId = int.parse(xfs[s].getAttribute('fillId')!);
      return fills[fillId].findAllElements('fgColor').single.getAttribute('rgb')!;
    }

    expect(fillOf('A1'), ExcelColor.black12.colorHex);
    expect(fillOf('A2'), ExcelColor.red.colorHex);
  });
}
