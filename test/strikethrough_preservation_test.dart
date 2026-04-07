import 'dart:io';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('Strikethrough Preservation Tests', () {
    test('Strikethrough style should be preserved in read/write cycle', () {
      // Create Excel with strikethrough text
      var excel = Excel.createExcel();
      var sheet = excel['Test'];
      
      // Cell A1: With strikethrough
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Strikethrough Text'),
        cellStyle: CellStyle(
          strikethrough: true,
        ),
      );
      
      // Cell A2: Without strikethrough
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('Normal Text'),
        cellStyle: CellStyle(
          strikethrough: false,
        ),
      );
      
      // Cell A3: Strikethrough + Bold
      sheet.updateCell(
        CellIndex.indexByString('A3'),
        TextCellValue('Strikethrough + Bold'),
        cellStyle: CellStyle(
          strikethrough: true,
          bold: true,
        ),
      );
      
      // Cell A4: Strikethrough + Italic
      sheet.updateCell(
        CellIndex.indexByString('A4'),
        TextCellValue('Strikethrough + Italic'),
        cellStyle: CellStyle(
          strikethrough: true,
          italic: true,
        ),
      );
      
      // Cell A5: Strikethrough + Bold + Italic + Underline
      sheet.updateCell(
        CellIndex.indexByString('A5'),
        TextCellValue('All Styles Combined'),
        cellStyle: CellStyle(
          strikethrough: true,
          bold: true,
          italic: true,
          underline: Underline.Single,
          fontColorHex: ExcelColor.red,
        ),
      );
      
      // Encode and decode
      var bytes = excel.encode()!;
      var excelDecoded = Excel.decodeBytes(bytes);
      var sheetDecoded = excelDecoded['Test'];
      
      // Verify A1 - Strikethrough
      var cell1 = sheetDecoded.cell(CellIndex.indexByString('A1'));
      expect(cell1.cellStyle?.isStrikethrough, equals(true),
          reason: 'A1 should have strikethrough');
      expect(cell1.value, equals(TextCellValue('Strikethrough Text')));
      
      // Verify A2 - No strikethrough
      var cell2 = sheetDecoded.cell(CellIndex.indexByString('A2'));
      expect(cell2.cellStyle?.isStrikethrough, equals(false),
          reason: 'A2 should NOT have strikethrough');
      
      // Verify A3 - Strikethrough + Bold
      var cell3 = sheetDecoded.cell(CellIndex.indexByString('A3'));
      expect(cell3.cellStyle?.isStrikethrough, equals(true),
          reason: 'A3 should have strikethrough');
      expect(cell3.cellStyle?.isBold, equals(true),
          reason: 'A3 should be bold');
      
      // Verify A4 - Strikethrough + Italic
      var cell4 = sheetDecoded.cell(CellIndex.indexByString('A4'));
      expect(cell4.cellStyle?.isStrikethrough, equals(true),
          reason: 'A4 should have strikethrough');
      expect(cell4.cellStyle?.isItalic, equals(true),
          reason: 'A4 should be italic');
      
      // Verify A5 - All styles combined
      var cell5 = sheetDecoded.cell(CellIndex.indexByString('A5'));
      expect(cell5.cellStyle?.isStrikethrough, equals(true),
          reason: 'A5 should have strikethrough');
      expect(cell5.cellStyle?.isBold, equals(true),
          reason: 'A5 should be bold');
      expect(cell5.cellStyle?.isItalic, equals(true),
          reason: 'A5 should be italic');
      expect(cell5.cellStyle?.underline, equals(Underline.Single),
          reason: 'A5 should have underline');
      expect(cell5.cellStyle?.fontColor, equals(ExcelColor.red),
          reason: 'A5 should have red font color');
    });
    
    test('Strikethrough should work with copyWith', () {
      var baseStyle = CellStyle(
        bold: true,
        fontSize: 14,
      );
      
      var strikethroughStyle = baseStyle.copyWith(strikethroughVal: true);
      
      expect(strikethroughStyle.isStrikethrough, equals(true));
      expect(strikethroughStyle.isBold, equals(true));
      expect(strikethroughStyle.fontSize, equals(14));
    });
    
    test('Strikethrough should be included in Equatable comparison', () {
      var style1 = CellStyle(strikethrough: true, bold: true);
      var style2 = CellStyle(strikethrough: true, bold: true);
      var style3 = CellStyle(strikethrough: false, bold: true);
      
      expect(style1, equals(style2),
          reason: 'Styles with same strikethrough should be equal');
      expect(style1, isNot(equals(style3)),
          reason: 'Styles with different strikethrough should not be equal');
    });
    
    test('Generate example file with strikethrough styles', () async {
      var excel = Excel.createExcel();
      var sheet = excel['Strikethrough Demo'];
      excel.delete('Sheet1');
      
      // Title
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('EJEMPLOS DE TACHADO (STRIKETHROUGH)'),
        cellStyle: CellStyle(
          bold: true,
          fontSize: 16,
          fontColorHex: ExcelColor.blue,
        ),
      );
      sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('C1'));
      
      int row = 3;
      
      // Sin tachado
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue('Sin tachado:'),
        cellStyle: CellStyle(bold: true),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        TextCellValue('Este texto NO está tachado'),
        cellStyle: CellStyle(strikethrough: false),
      );
      row++;
      
      // Con tachado
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue('Con tachado:'),
        cellStyle: CellStyle(bold: true),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        TextCellValue('Este texto está TACHADO'),
        cellStyle: CellStyle(
          strikethrough: true,
          fontColorHex: ExcelColor.red,
        ),
      );
      row++;
      
      // Tachado + Negrita
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue('Tachado + Negrita:'),
        cellStyle: CellStyle(bold: true),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        TextCellValue('Tachado y Negrita'),
        cellStyle: CellStyle(
          strikethrough: true,
          bold: true,
        ),
      );
      row++;
      
      // Tachado + Subrayado (interesante!)
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue('Tachado + Subrayado:'),
        cellStyle: CellStyle(bold: true),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        TextCellValue('Tachado y Subrayado'),
        cellStyle: CellStyle(
          strikethrough: true,
          underline: Underline.Single,
          fontColorHex: ExcelColor.purple,
        ),
      );
      row++;
      
      // Todos los estilos
      sheet.updateCell(
        CellIndex.indexByString('A$row'),
        TextCellValue('Todos los estilos:'),
        cellStyle: CellStyle(bold: true),
      );
      sheet.updateCell(
        CellIndex.indexByString('B$row'),
        TextCellValue('Negrita + Cursiva + Tachado + Subrayado'),
        cellStyle: CellStyle(
          strikethrough: true,
          bold: true,
          italic: true,
          underline: Underline.Double,
          fontColorHex: ExcelColor.green,
          fontSize: 14,
        ),
      );
      
      // Adjust column widths
      sheet.setColumnWidth(0, 25.0);
      sheet.setColumnWidth(1, 40.0);
      
      // Save file
      var bytes = excel.encode()!;
      var file = File('test/test_resources/strikethrough_example.xlsx');
      await file.create(recursive: true);
      await file.writeAsBytes(bytes);
      
      print('✅ Archivo generado: ${file.path}');
      print('📄 Tamaño: ${bytes.length} bytes');
      
      // Verify by reading back
      var excelRead = Excel.decodeBytes(bytes);
      var sheetRead = excelRead['Strikethrough Demo'];
      
      var cellWithStrike = sheetRead.cell(CellIndex.indexByString('B4'));
      expect(cellWithStrike.cellStyle?.isStrikethrough, equals(true),
          reason: 'Generated file should preserve strikethrough');
      
      print('✅ Strikethrough preservado correctamente!');
    });
  });
}
