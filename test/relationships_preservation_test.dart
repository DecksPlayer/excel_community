import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

// 1x1 PNGs (different pixels) so media files can be told apart.
final _red = base64Decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==');
final _blue = base64Decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==');

ExcelImage _image(List<int> bytes, int column) => ExcelImage(
      imageBytes: bytes,
      anchor: ImageAnchor.fromPixels(column: column, row: 1, widthPixels: 10, heightPixels: 10),
    );

Archive _zip(List<int> bytes) => ZipDecoder().decodeBytes(bytes);

String _part(Archive archive, String name) => utf8.decode(archive.findFile(name)!.content);

/// Relationship type (last path segment) for every Id in a .rels part.
Map<String, String> _relTypes(Archive archive, String relsPath) {
  final rels = _part(archive, relsPath);
  return {
    for (final m in RegExp(r'<Relationship [^>]*?Id="(rId\d+)"[^>]*?Type="[^"]*/(\w+)"')
        .allMatches(rels))
      m.group(1)!: m.group(2)!,
  };
}

void main() {
  test('adding a comment to a reopened file keeps its drawing relationship', () {
    final excel = Excel.createExcel();
    excel['Sheet1'].cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
    excel['Sheet1'].addImage(_image(_red, 1));

    final reopened = Excel.decodeBytes(excel.encode()!);
    reopened['Sheet1'].cell(CellIndex.indexByString('A1')).comment = 'note';
    final archive = _zip(reopened.encode()!);

    final types = _relTypes(archive, 'xl/worksheets/_rels/sheet1.xml.rels');
    expect(types.values, containsAll(['drawing', 'comments', 'vmlDrawing']));

    final sheetXml = _part(archive, 'xl/worksheets/sheet1.xml');
    final drawingId = RegExp(r'<drawing r:id="(rId\d+)"').firstMatch(sheetXml)!.group(1);
    final vmlId = RegExp(r'<legacyDrawing r:id="(rId\d+)"').firstMatch(sheetXml)!.group(1);
    expect(types[drawingId], 'drawing');
    expect(types[vmlId], 'vmlDrawing');
  });

  test('an image on another sheet does not overwrite the original drawing', () {
    final excel = Excel.createExcel();
    excel['Sheet1'].cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
    excel['Sheet1'].addImage(_image(_red, 1));
    final first = excel.encode()!;
    final originalDrawing = _part(_zip(first), 'xl/drawings/drawing1.xml');

    final reopened = Excel.decodeBytes(first);
    reopened['Second'].addImage(_image(_blue, 2));
    final archive = _zip(reopened.encode()!);

    expect(_part(archive, 'xl/drawings/drawing1.xml'), originalDrawing);
    expect(archive.findFile('xl/drawings/drawing2.xml'), isNotNull);
  });

  test('a new image in a reopened file does not overwrite existing media', () {
    final excel = Excel.createExcel();
    excel['Sheet1'].addImage(_image(_red, 1));

    final reopened = Excel.decodeBytes(excel.encode()!);
    reopened['Sheet1'].addImage(_image(_blue, 4));
    final archive = _zip(reopened.encode()!);

    expect(archive.findFile('xl/media/image1.png')!.content, _red);
    expect(archive.findFile('xl/media/image2.png')!.content, _blue);

    // Both images are anchored in the same drawing with distinct ids.
    final drawing = _part(archive, 'xl/drawings/drawing1.xml');
    expect(RegExp('<xdr:oneCellAnchor').allMatches(drawing).length, 2);
    final ids = _relTypes(archive, 'xl/drawings/_rels/drawing1.xml.rels').keys;
    expect(ids.toSet().length, 2);
  });

  test('new relationship ids skip gaps instead of colliding', () {
    final excel = Excel.createExcel();
    excel['Sheet1'].addImage(_image(_red, 1));
    final bytes = excel.encode()!;

    // Renumber the drawing relationship to rId2 (leaving a gap at rId1).
    final source = _zip(bytes);
    final patched = Archive();
    for (final file in source.files) {
      var content = file.content as List<int>;
      if (file.name == 'xl/worksheets/_rels/sheet1.xml.rels' ||
          file.name == 'xl/worksheets/sheet1.xml') {
        content = utf8.encode(utf8.decode(content).replaceAll('"rId1"', '"rId2"'));
      }
      patched.addFile(ArchiveFile(file.name, content.length, content));
    }

    final reopened = Excel.decodeBytes(ZipEncoder().encode(patched));
    reopened['Sheet1'].cell(CellIndex.indexByString('B2')).comment = 'note';
    final archive = _zip(reopened.encode()!);

    final types = _relTypes(archive, 'xl/worksheets/_rels/sheet1.xml.rels');
    expect(types['rId2'], 'drawing');
    expect(types.length, 3, reason: 'drawing + comments + vml, no duplicate ids');
  });
}
