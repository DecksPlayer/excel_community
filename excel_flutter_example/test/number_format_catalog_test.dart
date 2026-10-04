import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/data/number_format_catalog.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  test('library rendering matches what Excel shows', () {
    for (final entry in numFormatCatalog) {
      for (final sample in entry.samples) {
        expect(
          entry.format.format(sample.value).trim(),
          sample.expected,
          reason: '${entry.name} (${entry.format.formatCode})',
        );
      }
    }
  });

  test('every format accepts its sample values', () {
    for (final entry in numFormatCatalog) {
      for (final sample in entry.samples) {
        // Text keeps its explicit format even though accepts() is false.
        if (sample.value is TextCellValue) continue;
        expect(entry.format.accepts(sample.value), isTrue,
            reason: '${entry.name} should accept ${sample.value.runtimeType}');
      }
    }
  });

  test('covers every built-in format except the reserved IDs 23-26', () {
    final ids = numFormatCatalog.map((e) => e.numFmtId).whereType<int>().toList();
    final expected = [
      for (var id = 0; id <= 49; id++)
        if (id < 23 || id > 26) id,
    ];
    expect(ids.toSet(), expected.toSet());
    expect(ids.length, expected.length, reason: 'no duplicates');
  });

  test('snippets reference the format and its expected output', () {
    final entry = numFormatCatalog.firstWhere((e) => e.constant == 'standard_4');
    expect(
      entry.snippet,
      "final cell = sheet.cell(CellIndex.indexByString('A1'));\n"
      'cell.value = DoubleCellValue(1234.567);\n'
      'cell.cellStyle = CellStyle(\n'
      '  numberFormat: NumFormat.standard_4, // #,##0.00\n'
      ');\n'
      '// Excel shows: 1,234.57',
    );

    final euro = numFormatCatalog.firstWhere((e) => e.name == 'Euro Symbol');
    expect(euro.snippet, contains(r"NumFormat.custom(formatCode: r'[$€-2] #,##0.00')"));
  });
}
