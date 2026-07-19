import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:archive/archive.dart';

void main() {
  test('Create Excel with programmatic Pivot Table and verify XML/ZIP integrity', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sales Data'];
    excel.delete('Sheet1');

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Region"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"));
    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(100));

    var reportSheet = excel['Pivot Report'];
    final pivotTable = PivotTable(
      name: 'PivotTable1',
      sourceSheet: 'Sales Data',
      sourceRange: 'A1:B2',
      targetCell: CellIndex.indexByString('A3'),
      rows: ['Region'],
      values: [
        PivotTableValue(
          field: 'Amount',
          function: PivotValueFunction.sum,
        ),
      ],
    );

    reportSheet.addPivotTable(pivotTable);

    var bytes = excel.save();
    expect(bytes, isNotNull);

    // Verify ZIP integrity
    var archive = ZipDecoder().decodeBytes(bytes!);

    // Check files existence
    bool foundPivotTable =
        archive.files.any((f) => f.name == 'xl/pivotTables/pivotTable1.xml');
    bool foundPivotTableRels =
        archive.files.any((f) => f.name == 'xl/pivotTables/_rels/pivotTable1.xml.rels');
    bool foundPivotCacheDef =
        archive.files.any((f) => f.name == 'xl/pivotCache/pivotCacheDefinition1.xml');
    bool foundPivotCacheDefRels =
        archive.files.any((f) => f.name == 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels');
    bool foundPivotCacheRec =
        archive.files.any((f) => f.name == 'xl/pivotCache/pivotCacheRecords1.xml');

    expect(foundPivotTable, isTrue, reason: 'pivotTable1.xml must be present');
    expect(foundPivotTableRels, isTrue, reason: 'pivotTable1.xml.rels must be present');
    expect(foundPivotCacheDef, isTrue, reason: 'pivotCacheDefinition1.xml must be present');
    expect(foundPivotCacheDefRels, isTrue, reason: 'pivotCacheDefinition1.xml.rels must be present');
    expect(foundPivotCacheRec, isTrue, reason: 'pivotCacheRecords1.xml must be present');

    // Decode and ensure no errors
    var excel2 = Excel.decodeBytes(bytes);
    expect(excel2.sheets.containsKey('Pivot Report'), isTrue);
  });
}
