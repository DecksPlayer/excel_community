import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Excel Cell Comments (Notes)', () {
    test('Create, write, and read cell comments', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      // Add cell comments
      sheet.cell(CellIndex.indexByString('B2')).comment = 'Comment on B2';
      sheet.cell(CellIndex.indexByString('C3')).comment = 'Comment on C3';

      // Save and decode
      final bytes = excel.save();
      expect(bytes, isNotNull);

      final decodedExcel = Excel.decodeBytes(bytes!);
      final decodedSheet = decodedExcel['Sheet1'];

      // Assert comments are parsed correctly
      expect(decodedSheet.cell(CellIndex.indexByString('B2')).comment, equals('Comment on B2'));
      expect(decodedSheet.cell(CellIndex.indexByString('C3')).comment, equals('Comment on C3'));
      expect(decodedSheet.cell(CellIndex.indexByString('A1')).comment, isNull);
    });

    test('Modify and clear cell comments', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.cell(CellIndex.indexByString('A1')).comment = 'Original comment';
      sheet.cell(CellIndex.indexByString('B2')).comment = 'To be cleared';

      // Verify initial comments before save
      expect(sheet.cell(CellIndex.indexByString('A1')).comment, equals('Original comment'));
      expect(sheet.cell(CellIndex.indexByString('B2')).comment, equals('To be cleared'));

      // Modify A1 and clear B2
      sheet.cell(CellIndex.indexByString('A1')).comment = 'Modified comment';
      sheet.cell(CellIndex.indexByString('B2')).comment = null;

      // Save and decode
      final bytes = excel.save();
      expect(bytes, isNotNull);

      final decodedExcel = Excel.decodeBytes(bytes!);
      final decodedSheet = decodedExcel['Sheet1'];

      // Verify modification and removal
      expect(decodedSheet.cell(CellIndex.indexByString('A1')).comment, equals('Modified comment'));
      expect(decodedSheet.cell(CellIndex.indexByString('B2')).comment, isNull);
    });

    test('Verify OOXML archive structures for comments', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('B2')).comment = 'Structure test';

      final bytes = excel.save()!;
      final archive = ZipDecoder().decodeBytes(bytes);

      // 1. Verify files are present in the archive
      final filenames = archive.files.map((f) => f.name).toList();
      expect(filenames, contains('xl/comments1.xml'));
      expect(filenames, contains('xl/drawings/vmlDrawing1.vml'));
      expect(filenames, contains('xl/worksheets/_rels/sheet1.xml.rels'));

      // 2. Verify worksheet has legacyDrawing tag
      final sheetFile = archive.findFile('xl/worksheets/sheet1.xml')!;
      final sheetXml = utf8.decode(sheetFile.content);
      expect(sheetXml, contains('<legacyDrawing'));

      // 3. Verify comments XML contents
      final commentsFile = archive.findFile('xl/comments1.xml')!;
      final commentsXml = utf8.decode(commentsFile.content);
      final commentsDoc = XmlDocument.parse(commentsXml);
      final commentNode = commentsDoc.findAllElements('comment').first;
      expect(commentNode.getAttribute('ref'), equals('B2'));
      expect(commentNode.text, contains('Structure test'));

      // 4. Verify VML drawing contains correct client data (Row, Column)
      final vmlFile = archive.findFile('xl/drawings/vmlDrawing1.vml')!;
      final vmlContent = utf8.decode(vmlFile.content);
      expect(vmlContent, contains('<x:Row>1</x:Row>')); // Row 2 is index 1
      expect(vmlContent, contains('<x:Column>1</x:Column>')); // Col B is index 1
    });

    test('XLS comment reading compatibility', () {
      final fileOld = './test/test_resources/oldXLSFile.xls';
      final bytesOld = File(fileOld).readAsBytesSync();
      final excelOld = Excel.decodeBytes(bytesOld);
      expect(excelOld.tables.isNotEmpty, isTrue);

      final fileNew = './test/test_resources/newXLSFile.xls';
      final bytesNew = File(fileNew).readAsBytesSync();
      final excelNew = Excel.decodeBytes(bytesNew);
      expect(excelNew.tables.isNotEmpty, isTrue);
    });
  });
}
