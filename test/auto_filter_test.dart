import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('AutoFilter Model Tests', () {
    test('Construct AutoFilter with normalized coordinates', () {
      final filter = AutoFilter.fromRange(
        start: CellIndex.indexByString('A1'),
        end: CellIndex.indexByString('D10'),
      );

      expect(filter.ref, equals('A1:D10'));
      expect(filter.startColumn, equals(0));
      expect(filter.startRow, equals(0));
      expect(filter.endColumn, equals(3));
      expect(filter.endRow, equals(9));
      expect(filter.columnCount, equals(4));
      expect(filter.rowCount, equals(10));
      expect(filter.startCell, equals(CellIndex.indexByString('A1')));
      expect(filter.endCell, equals(CellIndex.indexByString('D10')));

      // Test containment
      expect(filter.containsCell(CellIndex.indexByString('B5')), isTrue);
      expect(filter.containsCell(CellIndex.indexByString('A1')), isTrue);
      expect(filter.containsCell(CellIndex.indexByString('D10')), isTrue);
      expect(filter.containsCell(CellIndex.indexByString('E5')), isFalse);
      expect(filter.containsCell(CellIndex.indexByString('B11')), isFalse);
      expect(filter.containsCellId('C3'), isTrue);
      expect(filter.containsCellId('Z99'), isFalse);
    });

    test('Inverted start and end coordinates are automatically normalized', () {
      final filter = AutoFilter.fromRange(
        start: CellIndex.indexByString('D10'),
        end: CellIndex.indexByString('A1'),
      );

      expect(filter.ref, equals('A1:D10'));
      expect(filter.startCell, equals(CellIndex.indexByString('A1')));
      expect(filter.endCell, equals(CellIndex.indexByString('D10')));
    });

    test('Construct AutoFilter with single cell range', () {
      final filter = AutoFilter.fromRangeString('B2');
      expect(filter.ref, equals('B2'));
      expect(filter.startCell, equals(CellIndex.indexByString('B2')));
      expect(filter.endCell, equals(CellIndex.indexByString('B2')));
      expect(filter.columnCount, equals(1));
      expect(filter.rowCount, equals(1));
      expect(filter.containsCellId('B2'), isTrue);
      expect(filter.containsCellId('B3'), isFalse);
    });

    test('XML serialization of simple AutoFilter', () {
      final filter = AutoFilter(ref: 'A1:C5');
      expect(filter.toXmlString(), equals('<autoFilter ref="A1:C5"/>'));
    });

    test('XML serialization with FilterColumn and filter values', () {
      final filter = AutoFilter(
        ref: 'A1:D20',
        filterColumns: [
          const FilterColumn(
            colId: 0,
            filterValues: ['Active', 'Pending'],
          ),
          const FilterColumn(
            colId: 2,
            blank: true,
          ),
        ],
      );

      final xml = filter.toXmlString();
      expect(xml, contains('<autoFilter ref="A1:D20">'));
      expect(xml, contains('<filterColumn colId="0"><filters><filter val="Active"/><filter val="Pending"/></filters></filterColumn>'));
      expect(xml, contains('<filterColumn colId="2"><filters blank="1"/></filterColumn>'));
      expect(xml, endsWith('</autoFilter>'));
    });

    test('XML serialization with CustomFilterRule', () {
      final filter = AutoFilter(
        ref: 'A1:B10',
        filterColumns: [
          const FilterColumn(
            colId: 1,
            customFilters: [
              CustomFilterRule(operator: FilterOperator.greaterThan, val: '50'),
            ],
          ),
        ],
      );

      final xml = filter.toXmlString();
      expect(xml, contains('<customFilters><customFilter operator="greaterThan" val="50"/></customFilters>'));
    });
  });

  group('Sheet AutoFilter API Tests', () {
    test('Set, modify, and clear AutoFilter on Sheet', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      expect(sheet.hasAutoFilter, isFalse);
      expect(sheet.autoFilter, isNull);

      sheet.setAutoFilter(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('F25'),
      );

      expect(sheet.hasAutoFilter, isTrue);
      expect(sheet.autoFilter?.ref, equals('A1:F25'));

      // Update by string
      sheet.setAutoFilterByString('B2:D12');
      expect(sheet.autoFilter?.ref, equals('B2:D12'));

      // Add filter column
      sheet.addFilterColumn(const FilterColumn(
        colId: 0,
        filterValues: ['Red', 'Blue'],
      ));
      expect(sheet.autoFilter?.filterColumns.length, equals(1));
      expect(sheet.autoFilter?.filterColumns.first.colId, equals(0));
      expect(sheet.autoFilter?.filterColumns.first.filterValues, equals(['Red', 'Blue']));

      // Clear
      sheet.clearAutoFilter();
      expect(sheet.hasAutoFilter, isFalse);
      expect(sheet.autoFilter, isNull);

      // Re-set and remove with alias
      sheet.setAutoFilterByString('A1:B2');
      expect(sheet.hasAutoFilter, isTrue);
      sheet.removeAutoFilter();
      expect(sheet.hasAutoFilter, isFalse);
    });
  });

  group('AutoFilter Save & Read (XLSX Round-Trip) Tests', () {
    test('Save and reload sheet with basic AutoFilter', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('ID');
      sheet.cell(CellIndex.indexByString('B1')).value = TextCellValue('Name');
      sheet.cell(CellIndex.indexByString('C1')).value = TextCellValue('Score');

      sheet.cell(CellIndex.indexByString('A2')).value = IntCellValue(1);
      sheet.cell(CellIndex.indexByString('B2')).value = TextCellValue('Alice');
      sheet.cell(CellIndex.indexByString('C2')).value = DoubleCellValue(95.5);

      sheet.cell(CellIndex.indexByString('A3')).value = IntCellValue(2);
      sheet.cell(CellIndex.indexByString('B3')).value = TextCellValue('Bob');
      sheet.cell(CellIndex.indexByString('C3')).value = DoubleCellValue(82.0);

      sheet.setAutoFilter(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('C3'),
      );

      final encoded = excel.save()!;
      expect(encoded.isNotEmpty, isTrue);

      // Verify raw XML in archive contains <autoFilter ref="A1:C3"/>
      final archive = ZipDecoder().decodeBytes(encoded);
      final sheetFile = archive.findFile('xl/worksheets/sheet1.xml')!;
      final sheetXml = utf8.decode(sheetFile.content);
      expect(sheetXml, contains('<autoFilter ref="A1:C3"/>'));

      // Decode with Excel
      final decodedExcel = Excel.decodeBytes(encoded);
      final decodedSheet = decodedExcel['Sheet1'];

      expect(decodedSheet.hasAutoFilter, isTrue);
      expect(decodedSheet.autoFilter?.ref, equals('A1:C3'));
      expect(decodedSheet.autoFilter?.startCell, equals(CellIndex.indexByString('A1')));
      expect(decodedSheet.autoFilter?.endCell, equals(CellIndex.indexByString('C3')));
      expect(decodedSheet.autoFilter?.columnCount, equals(3));
      expect(decodedSheet.autoFilter?.rowCount, equals(3));
    });

    test('Save and reload sheet with FilterColumn and values', () {
      final excel = Excel.createExcel();
      final sheet = excel['Inventory'];

      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Category');
      sheet.cell(CellIndex.indexByString('B1')).value = TextCellValue('Item');
      sheet.cell(CellIndex.indexByString('C1')).value = TextCellValue('Quantity');

      sheet.setAutoFilter(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('C10'),
      );

      sheet.addFilterColumn(const FilterColumn(
        colId: 0,
        filterValues: ['Tools', 'Hardware'],
      ));

      final encoded = excel.save()!;
      final decodedExcel = Excel.decodeBytes(encoded);
      final decodedSheet = decodedExcel['Inventory'];

      expect(decodedSheet.hasAutoFilter, isTrue);
      expect(decodedSheet.autoFilter?.ref, equals('A1:C10'));
      expect(decodedSheet.autoFilter?.filterColumns.length, equals(1));

      final col = decodedSheet.autoFilter!.filterColumns.first;
      expect(col.colId, equals(0));
      expect(col.filterValues, equals(['Tools', 'Hardware']));
      expect(col.blank, isFalse);
    });

    test('Save, clear AutoFilter, and reload removes autoFilter tag', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Col1');
      sheet.cell(CellIndex.indexByString('B1')).value = TextCellValue('Col2');
      sheet.setAutoFilterByString('A1:B10');

      // First save with autoFilter
      final encodedWithFilter = excel.save()!;
      final reloaded = Excel.decodeBytes(encodedWithFilter);
      expect(reloaded['Sheet1'].hasAutoFilter, isTrue);

      // Now clear it
      reloaded['Sheet1'].clearAutoFilter();
      final encodedWithoutFilter = reloaded.save()!;

      // Inspect archive XML
      final archive = ZipDecoder().decodeBytes(encodedWithoutFilter);
      final sheetXml = utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content);
      expect(sheetXml, isNot(contains('<autoFilter')));

      // Decode again
      final reloaded2 = Excel.decodeBytes(encodedWithoutFilter);
      expect(reloaded2['Sheet1'].hasAutoFilter, isFalse);
      expect(reloaded2['Sheet1'].autoFilter, isNull);
    });

    test('Schema compliance: autoFilter element ordering in worksheet XML', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Header1');
      sheet.cell(CellIndex.indexByString('B1')).value = TextCellValue('Header2');
      sheet.cell(CellIndex.indexByString('A2')).value = IntCellValue(10);
      sheet.cell(CellIndex.indexByString('B2')).value = IntCellValue(20);

      // Add autoFilter, mergeCells, and sheetProtection
      sheet.setAutoFilterByString('A1:B2');
      sheet.merge(CellIndex.indexByString('C1'), CellIndex.indexByString('D1'));
      sheet.protect('secret');

      final encoded = excel.save()!;
      final archive = ZipDecoder().decodeBytes(encoded);
      final sheetXml = utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content);

      // ECMA-376 schema order: sheetProtection -> autoFilter -> mergeCells
      final protectionIdx = sheetXml.indexOf('<sheetProtection');
      final autoFilterIdx = sheetXml.indexOf('<autoFilter');
      final mergeCellsIdx = sheetXml.indexOf('<mergeCells');

      expect(protectionIdx, isNonNegative);
      expect(autoFilterIdx, isNonNegative);
      expect(mergeCellsIdx, isNonNegative);

      expect(protectionIdx, lessThan(autoFilterIdx));
      expect(autoFilterIdx, lessThan(mergeCellsIdx));
    });
  });
}
