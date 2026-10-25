import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

/// The CT_Worksheet child sequence of ECMA-376 (Excel refuses a file whose
/// known elements are out of order).
const _schemaOrder = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData',
  'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter',
  'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
  'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
  'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects',
  'controls', 'webPublishItems', 'tableParts', 'extLst',
];

String _sheetXml(List<int> xlsx) => utf8.decode(
    ZipDecoder().decodeBytes(xlsx).findFile('xl/worksheets/sheet1.xml')!.content);

/// Top-level element names of a worksheet, in document order.
List<String> _children(String xml) {
  final body = xml.substring(xml.indexOf('>', xml.indexOf('<worksheet')) + 1);
  final names = <String>[];
  var depth = 0;
  for (final m in RegExp(r'<(/?)([\w:]+)[^>]*?(/?)>').allMatches(body)) {
    final closing = m.group(1)!.isNotEmpty;
    final selfClosing = m.group(3)!.isNotEmpty;
    if (closing) {
      depth--;
    } else {
      if (depth == 0 && m.group(2) != 'worksheet') names.add(m.group(2)!);
      if (!selfClosing) depth++;
    }
  }
  return names;
}

void expectSchemaOrder(List<String> children) {
  final known = children.where(_schemaOrder.contains).toList();
  final sorted = [...known]..sort((a, b) => _schemaOrder.indexOf(a).compareTo(_schemaOrder.indexOf(b)));
  expect(known, sorted, reason: 'worksheet children: $children');
}

void main() {
  test('re-saving keeps phoneticPr right after mergeCells', () {
    final bytes = File('test/test_resources/mergedBorders.xlsx').readAsBytesSync();
    final children = _children(_sheetXml(Excel.decodeBytes(bytes).encode()!));
    expect(children, containsAll(['mergeCells', 'phoneticPr', 'pageSetup']));
    expectSchemaOrder(children);
  });

  test('page breaks and ignored errors from a file stay in schema order', () {
    // A file with <rowBreaks>, <colBreaks> and <ignoredErrors>, as Excel
    // writes them; the library then adds its own elements around them.
    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];
    for (var r = 0; r < 5; r++) {
      sheet.appendRow([TextCellValue('R$r'), TextCellValue('${r * 10}')]);
    }
    final archive = ZipDecoder().decodeBytes(excel.encode()!);
    final xml = _sheetXml(excel.encode()!).replaceFirst(
        '</worksheet>',
        '<rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks>'
            '<colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks>'
            '<ignoredErrors><ignoredError sqref="B1:B5" numberStoredAsText="1"/></ignoredErrors>'
            '</worksheet>');
    final out = Archive();
    for (final f in archive.files) {
      final content = f.name == 'xl/worksheets/sheet1.xml' ? utf8.encode(xml) : f.content;
      out.addFile(ArchiveFile(f.name, content.length, content));
    }

    final reopened = Excel.decodeBytes(ZipEncoder().encode(out));
    reopened['Sheet1'].merge(CellIndex.indexByString('D1'), CellIndex.indexByString('E1'));
    reopened['Sheet1'].pageSetup = const PageSetup(orientation: PageOrientation.landscape);
    // A table makes the library write <tableParts>, which must come after them.
    reopened['Sheet1'].addTable('A1:B5', name: 'Data');
    final children = _children(_sheetXml(reopened.encode()!));
    expect(children,
        containsAll(['mergeCells', 'pageSetup', 'rowBreaks', 'colBreaks', 'ignoredErrors', 'tableParts']));
    expectSchemaOrder(children);
  });

  test('header and footer parts are written odd, even, first', () {
    final excel = Excel.createExcel();
    excel['Sheet1'].headerFooter = HeaderFooter(
      differentFirst: true,
      differentOddEven: true,
      firstHeader: '&CFirst',
      firstFooter: '&CFirst footer',
      evenHeader: '&CEven',
      evenFooter: '&CEven footer',
      oddHeader: '&COdd',
      oddFooter: '&COdd footer',
    );
    final xml = _sheetXml(excel.encode()!);
    final parts = RegExp(r'<(odd|even|first)(Header|Footer)>')
        .allMatches(xml)
        .map((m) => '${m.group(1)}${m.group(2)}')
        .toList();
    expect(parts, ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']);
  });
}
