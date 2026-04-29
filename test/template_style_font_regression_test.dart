import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  test('Template style edits preserve font color when styles.xml has extra font entries', () {
    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];
    final cellIndex = CellIndex.indexByString('B2');

    sheet.updateCell(
      cellIndex,
      TextCellValue('template'),
      cellStyle: CellStyle(
        fontColorHex: ExcelColor.white,
        backgroundColorHex: ExcelColor.black,
      ),
    );

    final templateBytes = _injectExtraFontEntry(excel.encode()!);
    final templateExcel = Excel.decodeBytes(templateBytes);

    final originalStyle = templateExcel['Sheet1'].cell(cellIndex).cellStyle!;
    expect(originalStyle.fontColor, ExcelColor.white);

    templateExcel['Sheet1'].updateCell(
      cellIndex,
      TextCellValue('updated'),
      cellStyle: originalStyle.copyWith(boldVal: true),
    );

    final updatedBytes = templateExcel.encode()!;
    final updatedExcel = Excel.decodeBytes(updatedBytes);
    final updatedStyle = updatedExcel['Sheet1'].cell(cellIndex).cellStyle!;

    expect(
      updatedStyle.fontColor,
      ExcelColor.white,
      reason: 'Template-derived styles should keep the original font color.',
    );

    final stylesArchive = ZipDecoder().decodeBytes(updatedBytes).findFile('xl/styles.xml')!;
    final stylesDocument = XmlDocument.parse(utf8.decode(stylesArchive.content));
    final fonts = stylesDocument.findAllElements('fonts').first;

    expect(
      fonts.getAttribute('count'),
      fonts.findElements('font').length.toString(),
      reason: 'styles.xml font count should match the actual number of font nodes.',
    );
  });
}

List<int> _injectExtraFontEntry(List<int> bytes) {
  final archive = ZipDecoder().decodeBytes(bytes);
  final stylesFile = archive.findFile('xl/styles.xml')!;
  final stylesDocument = XmlDocument.parse(utf8.decode(stylesFile.content));
  final fonts = stylesDocument.findAllElements('fonts').first;
  final fontNodes = fonts.findElements('font').toList();

  expect(fontNodes.length, greaterThanOrEqualTo(2));

  final duplicateDefaultFont = XmlDocument.parse(fontNodes.first.toXmlString())
      .rootElement
      .copy();
  fonts.children.insert(1, duplicateDefaultFont);
  fonts.setAttribute('count', fonts.findElements('font').length.toString());

  for (final xf in stylesDocument.findAllElements('cellXfs').first.findElements('xf')) {
    final fontId = int.tryParse(xf.getAttribute('fontId') ?? '0') ?? 0;
    if (fontId >= 1) {
      xf.setAttribute('fontId', (fontId + 1).toString());
    }
  }

  final updatedArchive = Archive();
  for (final file in archive.files) {
    final content = file.name == 'xl/styles.xml'
        ? utf8.encode(stylesDocument.toXmlString())
        : List<int>.from(file.content as List<int>);
    updatedArchive.addFile(ArchiveFile(file.name, content.length, content));
  }

  return ZipEncoder().encode(updatedArchive)!;
}