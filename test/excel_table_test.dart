import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

Archive _zip(List<int> bytes) => ZipDecoder().decodeBytes(bytes);
String _part(Archive a, String name) => utf8.decode(a.findFile(name)!.content);
CellIndex _at(String ref) => CellIndex.indexByString(ref);

Sheet _salesSheet(Excel excel) {
  final sheet = excel['Sheet1'];
  sheet.appendRow([TextCellValue('Region'), TextCellValue('Units'), TextCellValue('Price')]);
  sheet.appendRow([TextCellValue('North'), IntCellValue(10), DoubleCellValue(2.5)]);
  sheet.appendRow([TextCellValue('South'), IntCellValue(7), DoubleCellValue(3.0)]);
  return sheet;
}

void main() {
  group('Model', () {
    test('styles and structured references', () {
      expect(TableStyle.medium(9).name, 'TableStyleMedium9');
      expect(TableStyle.light(1).name, 'TableStyleLight1');
      expect(() => TableStyle.dark(12), throwsRangeError);
      const t = ExcelTable(name: 'Sales', ref: 'A1:C4', columns: [
        TableColumn('Region'), TableColumn('Unit [USD]'), TableColumn('#'),
      ]);
      expect(t.columnReference('Unit [USD]'), "Sales[Unit '[USD']]");
      expect(t.columnReference('#'), "Sales['#]");
      expect(t.firstDataRow, 1);
      expect(t.lastDataRow, 3);
      expect(t.columnIndexOf('unit [usd]'), 1);
    });
  });

  group('addTable', () {
    test('takes column names from the header row', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      final table = sheet.addTable('A1:C3', name: 'Sales');
      expect(table.columns.map((c) => c.name), ['Region', 'Units', 'Price']);
      expect(sheet.getTable('sales'), table);
      expect(sheet.tableAt(_at('B2'))?.name, 'Sales');
      expect(sheet.tableAt(_at('D2')), isNull);
    });

    test('empty and repeated headers get unique names', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Name'), null, TextCellValue('Name')]);
      sheet.appendRow([TextCellValue('a'), TextCellValue('b'), TextCellValue('c')]);
      final table = sheet.addTable('A1:C2', name: 'People');
      expect(table.columns.map((c) => c.name), ['Name', 'Column2', 'Name2']);
      expect(sheet.cell(_at('B1')).value, TextCellValue('Column2'), reason: 'header synced');
    });

    test('explicit columns overwrite the header cells and fill the totals row', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C4',
          name: 'Sales',
          showTotalsRow: true,
          columns: const [
            TableColumn('Region', totalsLabel: 'Total'),
            TableColumn('Qty', totalsFunction: TableTotalsFunction.sum),
            TableColumn('Price', totalsFunction: TableTotalsFunction.average),
          ]);
      expect(sheet.cell(_at('B1')).value, TextCellValue('Qty'));
      expect(sheet.cell(_at('A4')).value, TextCellValue('Total'));
      expect(sheet.cell(_at('B4')).value, FormulaCellValue('SUBTOTAL(109,Sales[Qty])'));
      expect(sheet.cell(_at('C4')).value, FormulaCellValue('SUBTOTAL(101,Sales[Price])'));
    });

    test('validation', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      expect(() => sheet.addTable('A1:C3', name: 'My Table'), throwsArgumentError);
      expect(() => sheet.addTable('A1:C3', name: 'A1'), throwsArgumentError);
      expect(() => sheet.addTable('A1:C3', name: 'R1C1'), throwsArgumentError);
      expect(() => sheet.addTable('A1:C1', name: 'Short'), throwsArgumentError, reason: 'no data row');
      expect(() => sheet.addTable('A1:C3', name: 'Cols', columns: const [TableColumn('a')]),
          throwsArgumentError);
      sheet.addTable('A1:C3', name: 'Sales');
      expect(() => excel['Other'].addTable('A1:B2', name: 'SALES'), throwsArgumentError,
          reason: 'names are unique in the workbook');
      expect(() => sheet.addTable('C3:D5', name: 'Overlap'), throwsArgumentError);

      sheet.merge(_at('F1'), _at('G1'));
      expect(() => sheet.addTable('F1:G3', name: 'Merged'), throwsArgumentError);
      sheet.setAutoFilterByString('I1:J5');
      expect(() => sheet.addTable('I1:J3', name: 'Filtered'), throwsArgumentError);
    });
  });

  group('Totals row', () {
    const columns = [
      TableColumn('Region', totalsLabel: 'Total'),
      TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
      TableColumn('Price'),
    ];

    test('addTable refuses a totals row that holds data', () {
      final sheet = _salesSheet(Excel.createExcel());
      expect(
          () => sheet.addTable('A1:C3', name: 'Sales', showTotalsRow: true, columns: columns),
          throwsArgumentError);
      expect(sheet.hasTables, isFalse);
      expect(sheet.cell(_at('A3')).value, TextCellValue('South'), reason: 'data kept');
      expect(sheet.cell(_at('C3')).value, DoubleCellValue(3.0));
    });

    test('a removed table can be added again over its own totals row', () {
      final sheet = _salesSheet(Excel.createExcel());
      sheet.addTable('A1:C4', name: 'Sales', showTotalsRow: true, columns: columns);
      sheet.removeTable('Sales');
      expect(sheet.addTable('A1:C4', name: 'Sales', showTotalsRow: true, columns: columns).ref,
          'A1:C4');
    });

    test('updateTable turning totals on adds a row below the table', () {
      final sheet = _salesSheet(Excel.createExcel());
      sheet.appendRow([TextCellValue('Below')]);
      final table = sheet.addTable('A1:C3', name: 'Sales', columns: columns);
      sheet.updateTable(table.copyWith(showTotalsRow: true));

      expect(sheet.getTable('Sales')!.ref, 'A1:C4');
      expect(sheet.cell(_at('A3')).value, TextCellValue('South'), reason: 'last data row kept');
      expect(sheet.cell(_at('A4')).value, TextCellValue('Total'));
      expect(sheet.cell(_at('A5')).value, TextCellValue('Below'), reason: 'shifted down');
    });

    test('updateTable turning totals off removes the totals row', () {
      final sheet = _salesSheet(Excel.createExcel());
      final table =
          sheet.addTable('A1:C4', name: 'Sales', showTotalsRow: true, columns: columns);
      sheet.updateTable(table.copyWith(showTotalsRow: false));

      expect(sheet.getTable('Sales')!.ref, 'A1:C3');
      expect(sheet.cell(_at('A4')).value, isNull);
      expect(sheet.tableRowsAsMaps('Sales').map((r) => r['Region']), ['North', 'South']);
    });
  });

  group('Data access and editing', () {
    test('tableRowsAsMaps and appendTableRow with a totals row', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C4', name: 'Sales', showTotalsRow: true, columns: const [
        TableColumn('Region', totalsLabel: 'Total'),
        TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
        TableColumn('Price'),
      ]);
      // A1:C4 includes the original row 4 slot as the totals row.
      final updated = sheet.appendTableRow('Sales', [TextCellValue('East'), IntCellValue(4), DoubleCellValue(1.5)]);
      expect(updated.ref, 'A1:C5');
      expect(sheet.cell(_at('A5')).value, TextCellValue('Total'), reason: 'totals row moved down');
      expect(sheet.tableRowsAsMaps('Sales').map((r) => r['Region']), ['North', 'South', 'East']);
    });

    test('appendTableRow without totals grows the range', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales');
      expect(sheet.appendTableRow('Sales', [TextCellValue('West')]).ref, 'A1:C4');
    });

    test('row and column edits move and resize tables', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales');
      sheet.insertRow(0);
      expect(sheet.getTable('Sales')!.ref, 'A2:C4');
      sheet.insertColumn(1); // inside the table
      final grown = sheet.getTable('Sales')!;
      expect(grown.ref, 'A2:D4');
      expect(grown.columns.map((c) => c.name), ['Region', 'Column4', 'Units', 'Price']);
      sheet.removeColumn(1);
      expect(sheet.getTable('Sales')!.columns.map((c) => c.name), ['Region', 'Units', 'Price']);
      sheet.removeRow(1); // header row
      expect(sheet.hasTables, isFalse);
    });

    test('removeTable keeps the cells', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales');
      sheet.removeTable('Sales');
      expect(sheet.hasTables, isFalse);
      expect(sheet.cell(_at('A2')).value, TextCellValue('North'));
    });
  });

  group('XLSX round-trip', () {
    test('writes the table part, relationship, content type and tableParts', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C4', name: 'Sales', style: TableStyle.medium(9), showTotalsRow: true,
          columns: const [
            TableColumn('Region', totalsLabel: 'Total'),
            TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
            TableColumn('Price'),
          ]);
      final archive = _zip(excel.encode()!);

      final table = _part(archive, 'xl/tables/table1.xml');
      expect(table, contains('id="1" name="Sales" displayName="Sales" ref="A1:C4" totalsRowCount="1">'));
      expect(table, contains('<autoFilter ref="A1:C3"/>'));
      expect(table, contains('<tableColumn id="2" name="Units" totalsRowFunction="sum"/>'));
      expect(table, contains('<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" '
          'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'));

      final sheetXml = _part(archive, 'xl/worksheets/sheet1.xml');
      final rId = RegExp(r'<tableParts count="1"><tablePart r:id="(rId\d+)"/></tableParts>')
          .firstMatch(sheetXml)!
          .group(1);
      expect(sheetXml.indexOf('<tableParts'), greaterThan(sheetXml.indexOf('<pageMargins')));
      expect(_part(archive, 'xl/worksheets/_rels/sheet1.xml.rels'),
          contains('Id="$rId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
              'Target="../tables/table1.xml"'));
      expect(_part(archive, '[Content_Types].xml'),
          contains('PartName="/xl/tables/table1.xml" '
              'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"'));
      expect(sheetXml, contains('<f>SUBTOTAL(109,Sales[Units])</f>'));
    });

    test('decodes what it encodes, on several sheets', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales', style: TableStyle.light(1), showFirstColumn: true,
          showRowStripes: false, showFilterButtons: false);
      final other = excel['People'];
      other.appendRow([TextCellValue('Name')]);
      other.appendRow([TextCellValue('Ana')]);
      other.addTable('A1:A2', name: 'People_List', style: null);

      final decoded = Excel.decodeBytes(excel.encode()!);
      expect(decoded['Sheet1'].tables, sheet.tables);
      expect(decoded['People'].tables, other.tables);
    });

    test('re-saving rebuilds parts without leftovers', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales');
      excel['Other'].appendRow([TextCellValue('X')]);
      excel['Other'].appendRow([IntCellValue(1)]);
      excel['Other'].addTable('A1:A2', name: 'Other');

      final reopened = Excel.decodeBytes(excel.encode()!);
      reopened['Sheet1'].removeTable('Sales');
      final archive = _zip(reopened.encode()!);
      final tableParts = archive.files.where((f) => f.name.startsWith('xl/tables/')).map((f) => f.name);
      expect(tableParts, ['xl/tables/table1.xml']);
      expect(_part(archive, 'xl/tables/table1.xml'), contains('name="Other"'));
      expect(_part(archive, '[Content_Types].xml').split('table+xml').length - 1, 1);
      expect(_part(archive, 'xl/worksheets/sheet1.xml'), isNot(contains('<tableParts')));

      // Saving twice gives the same parts.
      final again = _zip(reopened.encode()!);
      expect(again.files.where((f) => f.name.startsWith('xl/tables/')).length, 1);
    });

    test('copied sheets get unique table names on save', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.addTable('A1:C3', name: 'Sales');
      excel.copy('Sheet1', 'Copy');
      final decoded = Excel.decodeBytes(excel.encode()!);
      expect(decoded['Sheet1'].tables.single.name, 'Sales');
      expect(decoded['Copy'].tables.single.name, 'Sales2');
    });

    test('reads Excel-written tables with calculated columns', () {
      final excel = Excel.createExcel();
      final sheet = _salesSheet(excel);
      sheet.updateCell(_at('D1'), TextCellValue('Revenue'));
      final bytes = excel.encode()!;

      const tableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" '
          'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
          'xr:uid="{00000000-000C-0000-FFFF-FFFF00000000}" id="3" name="Table3" displayName="Revenue_Table" '
          'ref="A1:D3" totalsRowShown="0"><autoFilter ref="A1:D3"/><tableColumns count="4">'
          '<tableColumn id="1" name="Region"/><tableColumn id="2" name="Units"/>'
          '<tableColumn id="3" name="Price"/><tableColumn id="4" name="Revenue">'
          '<calculatedColumnFormula>Revenue_Table[[#This Row],[Units]]*Revenue_Table[[#This Row],[Price]]'
          '</calculatedColumnFormula></tableColumn></tableColumns>'
          '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
          'showRowStripes="1" showColumnStripes="0"/></table>';
      final source = _zip(bytes);
      final patched = Archive();
      for (final file in source.files) {
        var content = file.content as List<int>;
        if (file.name == 'xl/worksheets/sheet1.xml') {
          content = utf8.encode(utf8.decode(content).replaceFirst(
              '</worksheet>', '<tableParts count="1"><tablePart r:id="rId9"/></tableParts></worksheet>'));
        }
        patched.addFile(ArchiveFile(file.name, content.length, content));
      }
      final rels = utf8.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
          'Target="../tables/table3.xml"/></Relationships>');
      patched.addFile(ArchiveFile('xl/worksheets/_rels/sheet1.xml.rels', rels.length, rels));
      final tableBytes = utf8.encode(tableXml);
      patched.addFile(ArchiveFile('xl/tables/table3.xml', tableBytes.length, tableBytes));

      final reopened = Excel.decodeBytes(ZipEncoder().encode(patched));
      final table = reopened['Sheet1'].getTable('Revenue_Table')!;
      expect(table.ref, 'A1:D3');
      expect(table.columns.last.calculatedColumnFormula, startsWith('Revenue_Table[[#This Row],[Units]]'));

      final saved = _zip(reopened.encode()!);
      expect(saved.findFile('xl/tables/table3.xml'), isNull, reason: 'rebuilt as table1.xml');
      expect(_part(saved, 'xl/tables/table1.xml'), contains('<calculatedColumnFormula>'));
    });
  });
}
