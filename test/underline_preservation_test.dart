import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('Underline Preservation Tests', () {
    test('Single underline preservation after read/write', () {
      // Create Excel with single underline
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var cellStyle = CellStyle(
        underline: Underline.Single,
        fontSize: 14,
        bold: true,
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Single Underline'),
        cellStyle: cellStyle,
      );

      // Save to bytes
      var bytes = excel.encode()!;

      // Decode back
      var excel2 = Excel.decodeBytes(bytes);
      var cell2 = excel2['Sheet1'].cell(CellIndex.indexByString('A1'));

      print(
          'Single Underline - Original: ${cellStyle.underline}, Decoded: ${cell2.cellStyle?.underline}');

      expect(
        cell2.cellStyle?.underline,
        equals(Underline.Single),
        reason: 'Single underline should be preserved',
      );
    });

    test('Double underline preservation after read/write', () {
      // Create Excel with double underline
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var cellStyle = CellStyle(
        underline: Underline.Double,
        fontSize: 14,
        italic: true,
      );

      sheet.updateCell(
        CellIndex.indexByString('B1'),
        TextCellValue('Double Underline'),
        cellStyle: cellStyle,
      );

      // Save to bytes
      var bytes = excel.encode()!;

      // Decode back
      var excel2 = Excel.decodeBytes(bytes);
      var cell2 = excel2['Sheet1'].cell(CellIndex.indexByString('B1'));

      print(
          'Double Underline - Original: ${cellStyle.underline}, Decoded: ${cell2.cellStyle?.underline}');

      expect(
        cell2.cellStyle?.underline,
        equals(Underline.Double),
        reason: 'Double underline should be preserved',
      );
    });

    test('No underline preservation after read/write', () {
      // Create Excel with no underline
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var cellStyle = CellStyle(
        underline: Underline.None,
        fontSize: 14,
      );

      sheet.updateCell(
        CellIndex.indexByString('C1'),
        TextCellValue('No Underline'),
        cellStyle: cellStyle,
      );

      // Save to bytes
      var bytes = excel.encode()!;

      // Decode back
      var excel2 = Excel.decodeBytes(bytes);
      var cell2 = excel2['Sheet1'].cell(CellIndex.indexByString('C1'));

      print(
          'No Underline - Original: ${cellStyle.underline}, Decoded: ${cell2.cellStyle?.underline}');

      expect(
        cell2.cellStyle?.underline,
        equals(Underline.None),
        reason: 'No underline should be preserved',
      );
    });

    test('Multiple cells with different underline styles', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Create cells with different underline styles
      var styles = <String, Underline>{
        'A2': Underline.None,
        'B2': Underline.Single,
        'C2': Underline.Double,
      };

      styles.forEach((cellRef, underlineStyle) {
        sheet.updateCell(
          CellIndex.indexByString(cellRef),
          TextCellValue('Cell $cellRef'),
          cellStyle: CellStyle(underline: underlineStyle),
        );
      });

      // Save and reload
      var bytes = excel.encode()!;
      var excel2 = Excel.decodeBytes(bytes);

      print('\n=== VERIFICACIÓN DE MÚLTIPLES UNDERLINES ===\n');

      // Verify each cell
      styles.forEach((cellRef, originalUnderline) {
        var cell = excel2['Sheet1'].cell(CellIndex.indexByString(cellRef));
        var loadedUnderline = cell.cellStyle?.underline;

        print('Celda $cellRef: $originalUnderline -> $loadedUnderline');

        expect(
          loadedUnderline,
          equals(originalUnderline),
          reason: 'Underline in $cellRef should be preserved',
        );
      });

      print('\n✅ TODOS LOS UNDERLINES SE PRESERVAN CORRECTAMENTE\n');
    });
  });
}
