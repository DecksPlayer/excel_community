import 'dart:convert';
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

  test('pivot tables are saved computed: cells, cache records and layout', () {
    final excel = Excel.createExcel();
    final data = excel['Data'];
    data.appendRow([TextCellValue('Region'), TextCellValue('Product'), TextCellValue('Units')]);
    for (final row in [
      ['North', 'A', 10],
      ['South', 'B', 7],
      ['North', 'B', 4],
      ['West', 'A', 9],
    ]) {
      data.appendRow([TextCellValue(row[0] as String), TextCellValue(row[1] as String), IntCellValue(row[2] as int)]);
    }
    excel['Pivot'].addPivotTable(PivotTable(
      name: 'PT',
      sourceSheet: 'Data',
      sourceRange: 'A1:C5',
      targetCell: CellIndex.indexByString('A3'),
      rows: ['Region'],
      columns: ['Product'],
      values: [PivotTableValue(field: 'Units')],
    ));

    final bytes = excel.encode()!;
    final archive = ZipDecoder().decodeBytes(bytes);
    String part(String name) => utf8.decode(archive.findFile(name)!.content);

    final table = part('xl/pivotTables/pivotTable1.xml');
    expect(table, contains('<location ref="A3:D8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'));
    expect(table, contains('<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems>'));
    expect(table, contains('<colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>'));
    expect(part('xl/pivotCache/pivotCacheDefinition1.xml'), contains('<sharedItems count="3"><s v="North"/><s v="South"/><s v="West"/></sharedItems>'));
    expect(part('xl/pivotCache/pivotCacheRecords1.xml'), contains('<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4"><r><x v="0"/><x v="0"/><n v="10"/></r>'));

    final pivot = Excel.decodeBytes(bytes)['Pivot'];
    String? text(String cell) => pivot.cell(CellIndex.indexByString(cell)).value?.toString();
    expect([text('A3'), text('B3'), text('A4'), text('B4'), text('C4'), text('D4')],
        ['Sum of Units', 'Column Labels', 'Row Labels', 'A', 'B', 'Grand Total']);
    expect([text('A5'), text('B5'), text('C5'), text('D5')], ['North', '10', '4', '14']);
    expect([text('A7'), text('B7'), text('C7'), text('D7')], ['West', '9', null, '9']);
    expect([text('A8'), text('B8'), text('C8'), text('D8')], ['Grand Total', '19', '11', '30']);
  });

  test('several data fields add the virtual Values field to colFields', () {
    final excel = Excel.createExcel();
    final data = excel['Data'];
    data.appendRow([TextCellValue('Region'), TextCellValue('Product'), TextCellValue('Units'), TextCellValue('Price')]);
    data.appendRow([TextCellValue('North'), TextCellValue('A'), IntCellValue(10), DoubleCellValue(2.5)]);

    String pivotXml(PivotTable pivot) {
      final copy = Excel.decodeBytes(excel.encode()!);
      copy['Pivot'].addPivotTable(pivot);
      final file = ZipDecoder().decodeBytes(copy.encode()!).findFile('xl/pivotTables/pivotTable1.xml')!;
      return utf8.decode(file.content);
    }

    final twoValues = pivotXml(PivotTable(
      name: 'P1',
      sourceSheet: 'Data',
      sourceRange: 'A1:D2',
      targetCell: CellIndex.indexByString('A3'),
      rows: ['Region'],
      columns: ['Product'],
      values: [
        PivotTableValue(field: 'Units'),
        PivotTableValue(field: 'Price', function: PivotValueFunction.average),
      ],
    ));
    expect(twoValues, contains('<colFields count="2"><field x="1"/><field x="-2"/></colFields>'));

    final oneValue = pivotXml(PivotTable(
      name: 'P1',
      sourceSheet: 'Data',
      sourceRange: 'A1:D2',
      targetCell: CellIndex.indexByString('A3'),
      rows: ['Region'],
      values: [PivotTableValue(field: 'Units')],
    ));
    expect(oneValue, isNot(contains('<colFields')));
  });
}
