import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

/// `excel_features.xlsx` was saved by Excel 16: sheet "Data" has a table
/// (ExcelTable, A1:C6 + totals), a list validation with an input prompt,
/// hyperlinks, collapsed row/column groups, landscape A4 page setup, a tab
/// color and a comment; it also carries Excel's `xl/calcChain.xml`.
void main() {
  Excel load() => Excel.decodeBytes(
      File('test/test_resources/excel_features.xlsx').readAsBytesSync());

  test('reads the features of an Excel-authored file', () {
    final excel = load();
    final sheet = excel['Data'];

    final table = sheet.getTable('ExcelTable')!;
    expect(table.showTotalsRow, isTrue);
    expect(sheet.tableRowsAsMaps('ExcelTable'), isNotEmpty);
    expect(sheet.dataValidations, isNotEmpty);
    expect(sheet.hyperlinks, isNotEmpty);
    expect(sheet.rowGroups, isNotEmpty);
    expect(sheet.pageSetup?.orientation, PageOrientation.landscape);
  });

  test('drops the stale calcChain when saving a modified file', () {
    final excel = load();
    excel['Data'].appendTableRow('ExcelTable', [TextCellValue('Extra'), IntCellValue(1), IntCellValue(2)]);

    final archive = ZipDecoder().decodeBytes(excel.save()!);
    String part(String name) => utf8.decode(archive.findFile(name)!.content);

    expect(archive.findFile('xl/calcChain.xml'), isNull);
    expect(part('xl/_rels/workbook.xml.rels'), isNot(contains('calcChain')));
    expect(part('[Content_Types].xml'), isNot(contains('calcChain')));

    // The saved file still opens with its features.
    final reopened = Excel.decodeBytes(excel.save()!);
    expect(reopened['Data'].getTable('ExcelTable'), isNotNull);
  });
}
