import 'dart:convert';

import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

Sheet _people() {
  final sheet = Excel.createExcel()['Sheet1'];
  sheet.appendRow([
    TextCellValue('Name'),
    TextCellValue('Age'),
    TextCellValue('Active'),
    TextCellValue('Joined'),
    TextCellValue('Score'),
  ]);
  sheet.appendRow([
    TextCellValue('Ana'),
    IntCellValue(31),
    BoolCellValue(true),
    DateCellValue(year: 2024, month: 3, day: 5),
    DoubleCellValue(9.5),
  ]);
  sheet.appendRow([
    TextCellValue('Luis'),
    IntCellValue(28),
    BoolCellValue(false),
    DateCellValue(year: 2025, month: 1, day: 20),
    DoubleCellValue(7.25),
  ]);
  return sheet;
}

void main() {
  group('rowsAsMaps', () {
    test('uses the header row as keys and native Dart values', () {
      expect(_people().rowsAsMaps(), [
        {'Name': 'Ana', 'Age': 31, 'Active': true, 'Joined': DateTime.utc(2024, 3, 5), 'Score': 9.5},
        {'Name': 'Luis', 'Age': 28, 'Active': false, 'Joined': DateTime.utc(2025, 1, 20), 'Score': 7.25},
      ]);
    });

    test('names empty headers by column letter and suffixes duplicates', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Name'), null, TextCellValue('Name'), TextCellValue('  ')]);
      sheet.appendRow([TextCellValue('a'), TextCellValue('b'), TextCellValue('c'), TextCellValue('d')]);
      expect(sheet.rowsAsMaps().single.keys, ['Name', 'B', 'Name_2', 'D']);
    });

    test('skips empty rows unless asked to keep them', () {
      final sheet = _people();
      sheet.updateCell(CellIndex.indexByString('A5'), TextCellValue('Eva'));
      expect(sheet.rowsAsMaps().map((r) => r['Name']), ['Ana', 'Luis', 'Eva']);
      expect(sheet.rowsAsMaps(skipEmptyRows: false).map((r) => r['Name']),
          ['Ana', 'Luis', null, 'Eva']);
    });

    test('supports a header below title rows', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Quarterly report')]);
      sheet.appendRow([TextCellValue('Region'), TextCellValue('Sales')]);
      sheet.appendRow([TextCellValue('North'), IntCellValue(120)]);
      expect(sheet.rowsAsMaps(headerRow: 1), [
        {'Region': 'North', 'Sales': 120},
      ]);
      expect(() => sheet.rowsAsMaps(headerRow: -1), throwsRangeError);
    });

    test('display text mode applies number formats', () {
      final sheet = _people();
      sheet.cell(CellIndex.indexByString('E2')).cellStyle =
          CellStyle(numberFormat: NumFormat.standard_10);
      sheet.cell(CellIndex.indexByString('D2')).cellStyle =
          CellStyle(numberFormat: NumFormat.custom(formatCode: 'dd/mm/yyyy'));
      final first = sheet.rowsAsMaps(mode: ExportValueMode.displayText).first;
      expect(first['Score'], '950.00%');
      expect(first['Joined'], '05/03/2024');
      expect(first['Age'], '31');
    });

    test('formulas export their cached value', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Total'), TextCellValue('Pending')]);
      sheet.appendRow([
        FormulaCellValue('SUM(1,2)', cachedValue: IntCellValue(3)),
        FormulaCellValue('SUM(1,2)'),
      ]);
      expect(sheet.rowsAsMaps().single, {'Total': 3, 'Pending': null});
    });

    test('times export as Duration', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Start')]);
      sheet.appendRow([TimeCellValue(hour: 9, minute: 30, second: 0)]);
      expect(sheet.rowsAsMaps().single['Start'], const Duration(hours: 9, minutes: 30));
    });
  });

  group('rowsAsValues', () {
    test('returns the full grid including the header', () {
      final values = _people().rowsAsValues();
      expect(values.length, 3);
      expect(values.first, ['Name', 'Age', 'Active', 'Joined', 'Score']);
      expect(values[1][1], 31);
    });
  });

  group('toJson', () {
    test('encodes dates, times and numbers JSON-friendly', () {
      final sheet = _people();
      sheet.updateCell(CellIndex.indexByString('F1'), TextCellValue('Start'));
      sheet.updateCell(CellIndex.indexByString('F2'), TimeCellValue(hour: 9, minute: 5, second: 7));
      sheet.updateCell(
          CellIndex.indexByString('G1'), TextCellValue('Updated'));
      sheet.updateCell(CellIndex.indexByString('G2'),
          DateTimeCellValue(year: 2026, month: 10, day: 3, hour: 14, minute: 30));
      final rows = jsonDecode(sheet.toJson()) as List;
      expect(rows.first, {
        'Name': 'Ana',
        'Age': 31,
        'Active': true,
        'Joined': '2024-03-05',
        'Score': 9.5,
        'Start': '09:05:07',
        'Updated': '2026-10-03T14:30:00.000Z',
      });
    });

    test('pretty prints with an indent', () {
      expect(_people().toJson(indent: '  '), contains('\n  {\n    "Name": "Ana"'));
    });

    test('workbook export has one array per sheet', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].appendRow([TextCellValue('A')]);
      excel['Sheet1'].appendRow([IntCellValue(1)]);
      excel['Other'].appendRow([TextCellValue('B')]);
      excel['Other'].appendRow([IntCellValue(2)]);
      expect(jsonDecode(excel.toJson()), {
        'Sheet1': [{'A': 1}],
        'Other': [{'B': 2}],
      });
      expect(excel.toMaps()['Other'], [{'B': 2}]);
    });
  });

  group('toCsv', () {
    test('quotes separators, quotes and line breaks', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRow([TextCellValue('Name'), TextCellValue('Note')]);
      sheet.appendRow([TextCellValue('Smith, Ana'), TextCellValue('said "hi"')]);
      sheet.appendRow([TextCellValue('Luis'), TextCellValue('line1\nline2')]);
      expect(
        sheet.toCsv(),
        'Name,Note\r\n"Smith, Ana","said ""hi"""\r\nLuis,"line1\nline2"',
      );
    });

    test('uses display text and a custom separator', () {
      final sheet = _people();
      sheet.cell(CellIndex.indexByString('E2')).cellStyle =
          CellStyle(numberFormat: NumFormat.standard_2);
      final lines = sheet.toCsv(separator: ';', lineTerminator: '\n').split('\n');
      expect(lines[1], 'Ana;31;TRUE;03-05-24;9.50');
    });
  });

  group('appendRowsFromMaps', () {
    test('writes a header and typed cells on an empty sheet', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.appendRowsFromMaps([
        {'Name': 'Ana', 'Age': 31, 'Joined': DateTime.utc(2024, 3, 5)},
        {'Name': 'Luis', 'Age': 28, 'Start': const Duration(hours: 9)},
      ]);
      expect(sheet.rowsAsValues().first, ['Name', 'Age', 'Joined', 'Start']);
      expect(sheet.cell(CellIndex.indexByString('C2')).value,
          DateCellValue(year: 2024, month: 3, day: 5));
      expect(sheet.cell(CellIndex.indexByString('D3')).value,
          TimeCellValue(hour: 9, minute: 0, second: 0));
    });

    test('matches existing headers and adds new columns', () {
      final sheet = _people();
      sheet.appendRowsFromMaps([
        {'Age': 40, 'Name': 'Eva', 'City': 'Lima'},
      ]);
      expect(sheet.rowsAsValues().first.last, 'City');
      expect(sheet.rowsAsMaps().last,
          {'Name': 'Eva', 'Age': 40, 'Active': null, 'Joined': null, 'Score': null, 'City': 'Lima'});
    });

    test('round-trips through JSON and a saved file', () {
      final source = _people().rowsAsMaps();
      final excel = Excel.createExcel();
      excel['Sheet1'].appendRowsFromMaps(source);
      final decoded = Excel.decodeBytes(excel.encode()!);
      expect(decoded['Sheet1'].rowsAsMaps(), source);
    });
  });
}
