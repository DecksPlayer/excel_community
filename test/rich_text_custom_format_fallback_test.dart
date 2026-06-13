import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Rich Text Styles, Underlines, Custom Formats, & Fallbacks Tests', () {
    test('Rich text span formatting is fully preserved after write and read', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      final styledText = TextCellValue.span(
        TextSpan(
          text: 'Normal ',
          children: [
            TextSpan(
              text: 'Bold & Red',
              style: CellStyle(
                bold: true,
                fontColorHex: ExcelColor.red,
              ),
            ),
            TextSpan(text: ' and '),
            TextSpan(
              text: 'Italic, Double Underline & Arial 16',
              style: CellStyle(
                italic: true,
                underline: Underline.Double,
                fontFamily: 'Arial',
                fontSize: 16,
              ),
            ),
          ],
        ),
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        styledText,
      );

      // Save and reload
      final bytes = excel.encode()!;
      final excel2 = Excel.decodeBytes(bytes);
      final cell = excel2['Sheet1'].cell(CellIndex.indexByString('A1'));

      expect(cell.value, isA<TextCellValue>());
      final loadedValue = cell.value as TextCellValue;

      expect(loadedValue.value.text, isNull);
      expect(loadedValue.value.children, hasLength(4));

      // First child: Normal (no style)
      final child0 = loadedValue.value.children![0];
      expect(child0.text, equals('Normal '));

      // Second child: Bold & Red
      final child1 = loadedValue.value.children![1];
      expect(child1.text, equals('Bold & Red'));
      expect(child1.style?.isBold, isTrue);
      expect(child1.style?.fontColor, equals(ExcelColor.red));

      // Third child: and (no styles)
      final child2 = loadedValue.value.children![2];
      expect(child2.text, equals(' and '));

      // Fourth child: Italic, Double Underline & Arial 16
      final child3 = loadedValue.value.children![3];
      expect(child3.text, equals('Italic, Double Underline & Arial 16'));
      expect(child3.style?.isItalic, isTrue);
      expect(child3.style?.underline, equals(Underline.Double));
      expect(child3.style?.fontFamily, equals('Arial'));
      expect(child3.style?.fontSize, equals(16));
    });

    test('Correctly parses <u val="single"/> as single underline', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Underline Test'),
        cellStyle: CellStyle(underline: Underline.Single),
      );
      final bytes = excel.encode()!;

      // Modify the styles.xml inside the zip to use <u val="single"/>
      final archive = ZipDecoder().decodeBytes(bytes);
      final stylesFile = archive.findFile('xl/styles.xml')!;
      var xmlString = utf8.decode(stylesFile.content);

      // Replace <u/> with <u val="single"/>
      expect(xmlString.contains('<u/>'), isTrue, reason: 'Generated XML should contain <u/>');
      xmlString = xmlString.replaceAll('<u/>', '<u val="single"/>');
      
      final newArchive = Archive();
      for (final file in archive.files) {
        if (file.name == 'xl/styles.xml') {
          final updatedContent = utf8.encode(xmlString);
          newArchive.addFile(ArchiveFile(
            'xl/styles.xml',
            updatedContent.length,
            updatedContent,
          ));
        } else {
          newArchive.addFile(file);
        }
      }

      // Re-encode zip
      final newBytes = ZipEncoder().encode(newArchive)!;

      // Decode and verify
      final excel2 = Excel.decodeBytes(newBytes);
      final cell = excel2['Sheet1'].cell(CellIndex.indexByString('A1'));
      expect(cell.cellStyle?.underline, equals(Underline.Single),
          reason: 'Underline style <u val="single"/> should be parsed as Underline.Single');
    });

    test('Custom number format does not create duplicate entries in styles.xml', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      final customFormat = CustomNumericNumFormat(formatCode: r'#,##0.00 "USD"');
      final style = CellStyle(numberFormat: customFormat);

      // Apply the same custom format to multiple cells
      sheet.updateCell(CellIndex.indexByString('A1'), DoubleCellValue(100.5), cellStyle: style);
      sheet.updateCell(CellIndex.indexByString('A2'), DoubleCellValue(200.75), cellStyle: style);
      sheet.updateCell(CellIndex.indexByString('A3'), DoubleCellValue(300.0), cellStyle: style);

      final bytes = excel.encode()!;
      
      // Let's decode the bytes as a zip archive and inspect styles.xml directly
      final archive = ZipDecoder().decodeBytes(bytes);
      final stylesFile = archive.findFile('xl/styles.xml')!;
      final xmlString = utf8.decode(stylesFile.content);
      final document = XmlDocument.parse(xmlString);

      // There should only be ONE numFmt with the formatCode customFormat
      final numFmts = document.findAllElements('numFmt');
      final matchingFormats = numFmts.where((node) => node.getAttribute('formatCode') == customFormat.formatCode);
      expect(matchingFormats.length, equals(1), reason: 'Should only define the custom format once in the XML');
    });

    test('Safe format parsing fallbacks (never crash on malformed format data)', () {
      // Create instances of format classes
      const numericFmt = NumFormat.standard_1; // numeric
      const dateFmt = NumFormat.standard_14; // mm-dd-yy
      const timeFmt = NumFormat.standard_20; // h:mm

      // Valid cases
      expect(numericFmt.read('123'), equals(IntCellValue(123)));
      expect(numericFmt.read('123.45'), equals(DoubleCellValue(123.45)));

      // Malformed/Text fallback cases
      expect(numericFmt.read('abc'), equals(TextCellValue('abc')));
      expect(numericFmt.read(''), equals(TextCellValue('')));
      
      expect(dateFmt.read('abc'), equals(TextCellValue('abc')));
      expect(dateFmt.read(''), equals(TextCellValue('')));

      expect(timeFmt.read('abc'), equals(TextCellValue('abc')));
      expect(timeFmt.read(''), equals(TextCellValue('')));
    });
  });
}
