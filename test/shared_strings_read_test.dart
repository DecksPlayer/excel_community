import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

/// A workbook whose cells A1..A[n] point to the shared string items [items],
/// written as given.
List<int> _workbookWithSharedStrings(List<String> items) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var i = 0; i < items.length; i++) {
    sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i), TextCellValue('placeholder $i'));
  }
  final source = ZipDecoder().decodeBytes(excel.encode()!);
  final sst = utf8.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'count="${items.length}" uniqueCount="${items.length}">${items.join()}</sst>');
  final out = Archive();
  for (final file in source.files) {
    out.addFile(file.name == 'xl/sharedStrings.xml'
        ? ArchiveFile(file.name, sst.length, sst)
        : ArchiveFile(file.name, file.size, file.content));
  }
  return ZipEncoder().encode(out);
}

void main() {
  test('plain and rich shared strings read the same as before the fast path', () {
    final items = [
      '<si><t>Plain</t></si>',
      '<si><t xml:space="preserve">  spaced  </t></si>',
      '<si><t>Fish &amp; Chips &lt;3</t></si>',
      '<si><t>Line 1\r\nLine 2</t></si>',
      '<si><t/></si>',
      '<si><t></t></si>',
      '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> rest</t></r></si>',
      '<si><t>Kanji</t><rPh sb="0" eb="1"><t>kana</t></rPh></si>',
      '<si>\n  <t>Indented</t>\n</si>',
    ];
    final sheet = Excel.decodeBytes(_workbookWithSharedStrings(items))['Sheet1'];
    String? text(int row) => sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row))
        .value
        ?.toString();

    expect([for (var i = 0; i < items.length; i++) text(i)], [
      'Plain',
      '  spaced  ',
      'Fish & Chips <3',
      'Line 1\r\nLine 2', // kept as stored, like before
      '',
      '',
      'Bold rest',
      'Kanji',
      'Indented',
    ]);
    final rich = sheet.cell(CellIndex.indexByString('A7')).value as TextCellValue;
    expect(rich.value.children!.first.style?.isBold, isTrue);
  });

  test('re-saving keeps the shared strings it read', () {
    final items = ['<si><t>Fish &amp; Chips</t></si>', '<si><t xml:space="preserve"> a </t></si>'];
    final excel = Excel.decodeBytes(_workbookWithSharedStrings(items));
    final again = Excel.decodeBytes(excel.encode()!)['Sheet1'];
    expect(again.cell(CellIndex.indexByString('A1')).value, TextCellValue('Fish & Chips'));
    expect(again.cell(CellIndex.indexByString('A2')).value, TextCellValue(' a '));
  });
}
