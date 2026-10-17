import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

const _spreadsheetMl =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// A one-sheet workbook built from its XML parts. With [prefix], the
/// SpreadsheetML parts write every element with it, as the Open XML SDK does
/// (`<x:sheet>`); without, they declare the namespace as the default one.
List<int> _workbook({String? prefix}) {
  final x = prefix == null ? '' : '$prefix:';
  final namespace = prefix == null
      ? 'xmlns="$_spreadsheetMl"'
      : 'xmlns:$prefix="$_spreadsheetMl"';
  const strings = ['Item', 'Amount', 'Coffee'];
  final parts = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>',
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8"?>'
        '<${x}workbook $namespace xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15ac" '
        'xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac">'
        '<mc:AlternateContent><mc:Choice Requires="x15">'
        '<x15ac:absPath url="C:\\exports\\"/></mc:Choice></mc:AlternateContent>'
        '<${x}sheets><${x}sheet name="Sheet1" sheetId="1" r:id="rId1"/></${x}sheets>'
        '</${x}workbook>',
    'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8"?>'
        '<${x}styleSheet $namespace>'
        '<${x}fonts count="1"><${x}font><${x}sz val="11"/></${x}font></${x}fonts>'
        '<${x}fills count="1"><${x}fill><${x}patternFill patternType="none"/></${x}fill></${x}fills>'
        '<${x}borders count="1"><${x}border/></${x}borders>'
        '<${x}cellStyleXfs count="1"><${x}xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></${x}cellStyleXfs>'
        '<${x}cellXfs count="2">'
        '<${x}xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<${x}xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '</${x}cellXfs>'
        '</${x}styleSheet>',
    'xl/sharedStrings.xml': '<?xml version="1.0" encoding="UTF-8"?>'
        '<${x}sst $namespace count="${strings.length}" uniqueCount="${strings.length}">'
        '${strings.map((s) => '<${x}si><${x}t>$s</${x}t></${x}si>').join()}'
        '</${x}sst>',
    'xl/worksheets/sheet1.xml': '<?xml version="1.0" encoding="UTF-8"?>'
        '<${x}worksheet $namespace><${x}sheetData>'
        '<${x}row r="1">'
        '<${x}c r="A1" t="s"><${x}v>0</${x}v></${x}c>'
        '<${x}c r="B1" t="s"><${x}v>1</${x}v></${x}c>'
        '</${x}row>'
        '<${x}row r="2">'
        '<${x}c r="A2" t="s"><${x}v>2</${x}v></${x}c>'
        '<${x}c r="B2" s="1"><${x}v>-2.5</${x}v></${x}c>'
        '</${x}row>'
        '</${x}sheetData></${x}worksheet>',
  };
  final archive = Archive();
  parts.forEach((name, content) =>
      archive.addFile(ArchiveFile.bytes(name, utf8.encode(content))));
  return ZipEncoder().encodeBytes(archive);
}

List<List<CellValue?>> _values(Excel excel) => [
      for (final row in excel.tables['Sheet1']!.rows)
        [for (final cell in row) cell?.value]
    ];

XmlDocument _part(List<int> xlsx, String name) => XmlDocument.parse(
    utf8.decode(ZipDecoder().decodeBytes(xlsx).findFile(name)!.content));

void main() {
  group('Workbooks whose parts use a namespace prefix', () {
    test('read as the same workbook written without the prefix', () {
      final prefixed = Excel.decodeBytes(_workbook(prefix: 'x'));
      final plain = Excel.decodeBytes(_workbook());

      expect(prefixed.tables.keys, ['Sheet1']);
      expect(_values(prefixed), _values(plain));
      expect(_values(prefixed)[1],
          [TextCellValue('Coffee'), DoubleCellValue(-2.5)]);
    });

    test('any prefix: it is the namespace that counts', () {
      final excel = Excel.decodeBytes(_workbook(prefix: 'ss'));
      expect(_values(excel), _values(Excel.decodeBytes(_workbook())));
    });

    test('encode as a regular workbook, edits included', () {
      final excel = Excel.decodeBytes(_workbook(prefix: 'x'));
      excel['Sheet1']
          .updateCell(CellIndex.indexByString('A3'), TextCellValue('Tea'));

      final bytes = excel.encode()!;
      final reread = Excel.decodeBytes(bytes);
      expect(_values(reread), _values(excel));
      expect(_values(reread)[2].first, TextCellValue('Tea'));

      final sheet = _part(bytes, 'xl/worksheets/sheet1.xml');
      expect(
        sheet.findAllElements('c').map((cell) => cell.name.qualified).toSet(),
        {'c'},
      );
      expect(sheet.rootElement.getAttribute('xmlns'), _spreadsheetMl);
    });

    test('keep the markup of other namespaces as written', () {
      final bytes = Excel.decodeBytes(_workbook(prefix: 'x')).encode()!;
      final workbook = _part(bytes, 'xl/workbook.xml').rootElement;

      expect(workbook.name.qualified, 'workbook');
      expect(workbook.getAttribute('mc:Ignorable'), 'x15ac');
      expect(workbook.findAllElements('mc:AlternateContent'), hasLength(1));
      expect(
          workbook.findAllElements('x15ac:absPath').single.getAttribute('url'),
          r'C:\exports\');
      expect(workbook.findAllElements('sheet').single.getAttribute('r:id'),
          'rId1');
    });

    test('workbooks written without a prefix are read as before', () {
      final bytes = _workbook();
      final excel = Excel.decodeBytes(bytes);
      expect(
          _values(excel)[1], [TextCellValue('Coffee'), DoubleCellValue(-2.5)]);
    });
  });
}
