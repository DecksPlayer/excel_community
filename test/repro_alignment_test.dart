import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  test('Reproduction: Alignment format loss after read/write', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];
    
    var cellIndex = CellIndex.indexByString('A1');
    var cellStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 45,
    );
    
    sheet.updateCell(cellIndex, TextCellValue('Aligned'), cellStyle: cellStyle);
    
    // Save to bytes
    var bytes = excel.encode()!;
    
    // Decode back
    var excel2 = Excel.decodeBytes(bytes);
    var cell2 = excel2['Sheet1'].cell(cellIndex);
    
    print('Original Alignment: H=${cellStyle.horizontalAlignment}, V=${cellStyle.verticalAlignment}, R=${cellStyle.rotation}');
    print('Decoded Alignment: H=${cell2.cellStyle?.horizontalAlignment}, V=${cell2.cellStyle?.verticalAlignment}, R=${cell2.cellStyle?.rotation}');
    
    expect(cell2.cellStyle?.horizontalAlignment, equals(HorizontalAlign.Center));
    expect(cell2.cellStyle?.verticalAlignment, equals(VerticalAlign.Center));
    expect(cell2.cellStyle?.rotation, equals(45));
  });
}
