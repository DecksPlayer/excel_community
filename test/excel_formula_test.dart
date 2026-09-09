import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

/// Loads a known-good fixture workbook and injects a formula cell with a
/// cached `<v>` result directly into its first worksheet's XML, mimicking
/// what Microsoft Excel writes for a calculated formula cell. This lets the
/// test exercise the real parser without hand-building an entire XLSX
/// package from scratch.
List<int> _fixtureWithFormulaCell(String cellXml) {
  final bytes = File('./test/test_resources/example.xlsx').readAsBytesSync();
  final archive = ZipDecoder().decodeBytes(bytes);

  final sheetFile = archive.findFile('xl/worksheets/sheet1.xml')!;
  sheetFile.decompress();
  final original = utf8.decode(sheetFile.content);
  final updated = original.replaceFirst(
    '</sheetData>',
    '<row r="10">$cellXml</row></sheetData>',
  );

  final content = utf8.encode(updated);
  final newArchive = Archive();
  for (final file in archive.files) {
    if (file.name == 'xl/worksheets/sheet1.xml') {
      newArchive.addFile(ArchiveFile(file.name, content.length, content));
    } else {
      file.decompress();
      newArchive
          .addFile(ArchiveFile(file.name, file.content.length, file.content));
    }
  }
  return ZipEncoder().encode(newArchive);
}

void main() {
  group('FormulaCellValue cached <v>', () {
    test('retains the pre-calculated <v> alongside the formula', () {
      final bytes =
          _fixtureWithFormulaCell('<c r="A10"><f>SUM(1,2)</f><v>3</v></c>');
      final excel = Excel.decodeBytes(bytes);
      final cell = excel.tables['Sheet1']!.rows[9][0]!.value;

      expect(cell, isA<FormulaCellValue>());
      final formula = cell as FormulaCellValue;
      expect(formula.formula, equals('SUM(1,2)'));
      expect(formula.cachedValue, equals(IntCellValue(3)));
    });

    test('cachedValue is null when the file has no <v>', () {
      final bytes = _fixtureWithFormulaCell('<c r="A10"><f>SUM(1,2)</f></c>');
      final excel = Excel.decodeBytes(bytes);
      final cell = excel.tables['Sheet1']!.rows[9][0]!.value;

      expect(cell, isA<FormulaCellValue>());
      expect((cell as FormulaCellValue).cachedValue, isNull);
    });

    test('cachedValue reflects a text-returning formula result', () {
      final bytes =
          _fixtureWithFormulaCell('<c r="A10"><f>A1&amp;"!"</f><v>3.5</v></c>');
      final excel = Excel.decodeBytes(bytes);
      final cell = excel.tables['Sheet1']!.rows[9][0]!.value;

      expect(
          (cell as FormulaCellValue).cachedValue, equals(DoubleCellValue(3.5)));
    });

    test('two FormulaCellValues with different cachedValue are not equal', () {
      const a = FormulaCellValue('SUM(1,2)', cachedValue: IntCellValue(3));
      const b = FormulaCellValue('SUM(1,2)', cachedValue: IntCellValue(4));
      const c = FormulaCellValue('SUM(1,2)', cachedValue: IntCellValue(3));

      expect(a == b, isFalse);
      expect(a == c, isTrue);
      expect(a.hashCode == c.hashCode, isTrue);
    });
  });
}
