import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

String _sheetXml(List<int> bytes, [String path = 'xl/worksheets/sheet1.xml']) {
  final archive = ZipDecoder().decodeBytes(bytes);
  return utf8.decode(archive.findFile(path)!.content);
}

/// Re-zips [bytes] with sheet1.xml rewritten by [transform].
List<int> _patchSheetXml(List<int> bytes, String Function(String) transform) {
  final archive = ZipDecoder().decodeBytes(bytes);
  final out = Archive();
  for (final file in archive.files) {
    if (file.name == 'xl/worksheets/sheet1.xml') {
      final patched = utf8.encode(transform(utf8.decode(file.content)));
      out.addFile(ArchiveFile(file.name, patched.length, patched));
    } else {
      out.addFile(ArchiveFile(file.name, file.size, file.content));
    }
  }
  return ZipEncoder().encode(out);
}

void main() {
  group('PageSetup model', () {
    test('toXmlString writes only set attributes in schema order', () {
      const setup = PageSetup(
        orientation: PageOrientation.landscape,
        paperSize: PaperSize.a4,
        scale: 85,
        copies: 2,
      );
      expect(
        setup.toXmlString(),
        equals('<pageSetup paperSize="9" scale="85" '
            'orientation="landscape" copies="2"/>'),
      );
    });

    test('toXmlString is empty when nothing is set', () {
      expect(const PageSetup().toXmlString(), isEmpty);
      expect(const PageSetup(fitToPage: true).toXmlString(), isEmpty);
    });

    test('toXmlString keeps a relationship id', () {
      expect(const PageSetup().toXmlString(relationshipId: 'rId1'),
          equals('<pageSetup r:id="rId1"/>'));
    });

    test('scale is clamped to 10-400 on write', () {
      expect(const PageSetup(scale: 5).toXmlString(), contains('scale="10"'));
      expect(
          const PageSetup(scale: 999).toXmlString(), contains('scale="400"'));
    });

    test('enum xml values', () {
      const setup = PageSetup(
        orientation: PageOrientation.automatic,
        pageOrder: PageOrder.overThenDown,
        cellComments: PrintCellComments.atEnd,
        errors: PrintErrors.na,
        blackAndWhite: true,
        draft: false,
      );
      final xml = setup.toXmlString();
      expect(xml, contains('orientation="default"'));
      expect(xml, contains('pageOrder="overThenDown"'));
      expect(xml, contains('cellComments="atEnd"'));
      expect(xml, contains('errors="NA"'));
      expect(xml, contains('blackAndWhite="1"'));
      expect(xml, contains('draft="0"'));
    });

    test('PaperSize.fromCode returns known constants or keeps raw codes', () {
      expect(PaperSize.fromCode(9), same(PaperSize.a4));
      expect(PaperSize.fromCode(1), same(PaperSize.letter));
      final custom = PaperSize.fromCode(117);
      expect(custom.code, equals(117));
      expect(custom.widthMm, isNull);
      expect(custom, equals(PaperSize.fromCode(117)));
    });

    test('copyWith and equality', () {
      const a = PageSetup(orientation: PageOrientation.portrait);
      final b = a.copyWith(paperSize: PaperSize.legal);
      expect(b.orientation, PageOrientation.portrait);
      expect(b.paperSize, PaperSize.legal);
      expect(a == b, isFalse);
      expect(b, equals(a.copyWith(paperSize: PaperSize.legal)));
    });
  });

  group('PageMargins & PrintOptions models', () {
    test('default margins match Excel "Normal" preset', () {
      expect(
        PageMargins.normal.toXmlString(),
        equals('<pageMargins left="0.7" right="0.7" top="0.75" '
            'bottom="0.75" header="0.3" footer="0.3"/>'),
      );
    });

    test('presets and all()', () {
      expect(PageMargins.wide.left, 1);
      expect(PageMargins.narrow.left, 0.25);
      expect(PageMargins.narrow.top, 0.75);
      const all = PageMargins.all(0.5);
      expect([all.left, all.right, all.top, all.bottom], everyElement(0.5));
      expect(all.toXmlString(), contains('left="0.5"'));
    });

    test('fromCentimeters converts to inches', () {
      final m = PageMargins.fromCentimeters(left: 2.54, right: 5.08);
      expect(m.left, closeTo(1.0, 1e-9));
      expect(m.right, closeTo(2.0, 1e-9));
    });

    test('PrintOptions serialization', () {
      expect(const PrintOptions().toXmlString(), isEmpty);
      expect(
        const PrintOptions(
          gridLines: true,
          headings: true,
          horizontalCentered: true,
        ).toXmlString(),
        equals('<printOptions horizontalCentered="1" headings="1" '
            'gridLines="1"/>'),
      );
    });
  });

  group('Sheet page setup API', () {
    test('convenience setters merge into the existing page setup', () {
      final sheet = Excel.createExcel()['Sheet1'];
      expect(sheet.hasPageSetup, isFalse);

      sheet.setPageOrientation(PageOrientation.landscape);
      sheet.setPaperSize(PaperSize.a3);
      expect(sheet.pageSetup!.orientation, PageOrientation.landscape);
      expect(sheet.pageSetup!.paperSize, PaperSize.a3);

      sheet.fitToPages(width: 1, height: 0);
      expect(sheet.pageSetup!.fitToPage, isTrue);
      expect(sheet.pageSetup!.fitToWidth, 1);
      expect(sheet.pageSetup!.fitToHeight, 0);

      sheet.setPrintScale(75);
      expect(sheet.pageSetup!.scale, 75);
      expect(sheet.pageSetup!.fitToPage, isFalse);
      expect(sheet.pageSetup!.orientation, PageOrientation.landscape);
    });

    test('validation', () {
      final sheet = Excel.createExcel()['Sheet1'];
      expect(() => sheet.setPrintScale(5), throwsRangeError);
      expect(() => sheet.setPrintScale(401), throwsRangeError);
      expect(() => sheet.fitToPages(width: -1), throwsRangeError);
      expect(() => sheet.pageMargins = const PageMargins(left: -0.1),
          throwsArgumentError);
    });

    test('clear and remove', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.setPageOrientation(PageOrientation.portrait);
      sheet.setPrintGridLines(true);
      sheet.pageMargins = PageMargins.wide;

      sheet.removePageSetup();
      sheet.clearPrintOptions();
      sheet.clearPageMargins();
      expect(sheet.hasPageSetup, isFalse);
      expect(sheet.hasPrintOptions, isFalse);
      expect(sheet.pageMargins, isNull);
    });

    test('setting empty PrintOptions clears them', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.printOptions = const PrintOptions();
      expect(sheet.hasPrintOptions, isFalse);
    });

    test('copying a sheet preserves print configuration', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setPageOrientation(PageOrientation.landscape);
      sheet.pageMargins = PageMargins.narrow;
      sheet.setPrintHeadings(true);

      excel.copy('Sheet1', 'Copy');
      final copy = excel['Copy'];
      expect(copy.pageSetup, equals(sheet.pageSetup));
      expect(copy.pageMargins, equals(PageMargins.narrow));
      expect(copy.printOptions!.headings, isTrue);
    });

    test('page setup on default Sheet1 prevents it being renamed', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setPageOrientation(PageOrientation.landscape);
      excel['Report'];
      expect(excel.tables.keys, containsAll(['Sheet1', 'Report']));
    });
  });

  group('XLSX page setup round-trip & schema order', () {
    test('writes elements in CT_Worksheet order', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
      sheet.setPageOrientation(PageOrientation.landscape);
      sheet.setPaperSize(PaperSize.a4);
      sheet.fitToPages(width: 1, height: 0);
      sheet.pageMargins = PageMargins.wide;
      sheet.setPrintGridLines(true);
      sheet.headerFooter = HeaderFooter(oddHeader: '&CReport');
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);

      final xml = _sheetXml(excel.encode()!);

      expect(xml, contains('<pageSetUpPr fitToPage="1"/>'));
      expect(
          xml,
          contains('<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" '
              'orientation="landscape"/>'));

      final sheetPr = xml.indexOf('<sheetPr');
      final sheetData = xml.indexOf('<sheetData');
      final printOptions = xml.indexOf('<printOptions');
      final pageMargins = xml.indexOf('<pageMargins');
      final pageSetup = xml.indexOf('<pageSetup ');
      final headerFooter = xml.indexOf('<headerFooter');
      expect(sheetPr, lessThan(sheetData));
      expect(sheetData, lessThan(printOptions));
      expect(printOptions, lessThan(pageMargins));
      expect(pageMargins, lessThan(pageSetup));
      expect(pageSetup, lessThan(headerFooter));

      // CT_SheetPr: outlinePr before pageSetUpPr.
      expect(xml.indexOf('<outlinePr'), isNonNegative);
      expect(xml.indexOf('<outlinePr'), lessThan(xml.indexOf('<pageSetUpPr')));
      expect('<pageMargins'.allMatches(xml).length, 1);
    });

    test('sheetPr ordering with tab color', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.setTabColorHex('#FF0000');
      sheet.fitToPages();
      sheet.outlineSettings =
          const OutlineSettings(summaryBelow: false, summaryRight: false);
      final xml = _sheetXml(excel.encode()!);
      expect(
          xml,
          contains('<sheetPr><tabColor rgb="FFFF0000"/>'
              '<outlinePr summaryBelow="0" summaryRight="0"/>'
              '<pageSetUpPr fitToPage="1"/></sheetPr>'));
    });

    test('decodes what it encodes', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
      sheet.pageSetup = PageSetup(
        orientation: PageOrientation.landscape,
        paperSize: PaperSize.legal,
        fitToPage: true,
        fitToWidth: 2,
        fitToHeight: 3,
        firstPageNumber: 5,
        useFirstPageNumber: true,
        pageOrder: PageOrder.overThenDown,
        blackAndWhite: true,
        draft: true,
        cellComments: PrintCellComments.asDisplayed,
        errors: PrintErrors.dash,
        horizontalDpi: 300,
        verticalDpi: 300,
        copies: 3,
      );
      sheet.pageMargins = const PageMargins(
          left: 0.5, right: 0.6, top: 1, bottom: 1.2, header: 0.4, footer: 0.45);
      sheet.printOptions = const PrintOptions(
          gridLines: true,
          headings: true,
          horizontalCentered: true,
          verticalCentered: true);

      final decoded = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(decoded.pageSetup, equals(sheet.pageSetup));
      expect(decoded.pageMargins, equals(sheet.pageMargins));
      expect(decoded.printOptions, equals(sheet.printOptions));
    });

    test('settings on newly created sheets survive save', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(CellIndex.indexByString('A1')).value =
          TextCellValue('keep Sheet1');
      final sheet = excel['Invoice'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('x');
      sheet.pageMargins = PageMargins.narrow;
      sheet.setPaperSize(PaperSize.letter);
      sheet.setPrintGridLines(true);

      final decoded = Excel.decodeBytes(excel.encode()!)['Invoice'];
      expect(decoded.pageMargins, equals(PageMargins.narrow));
      expect(decoded.pageSetup!.paperSize, PaperSize.letter);
      expect(decoded.printOptions!.gridLines, isTrue);
    });

    test('new sheets get default margins and no pageSetup', () {
      final excel = Excel.createExcel();
      excel['Data'].cell(CellIndex.indexByString('A1')).value =
          TextCellValue('x');
      final bytes = excel.encode()!;
      final archive = ZipDecoder().decodeBytes(bytes);
      for (final f in archive.files
          .where((f) => f.name.startsWith('xl/worksheets/sheet'))) {
        final xml = utf8.decode(f.content);
        expect('<pageMargins'.allMatches(xml).length, 1, reason: f.name);
        expect(xml, contains(PageMargins.normal.toXmlString()));
        expect(xml, isNot(contains('<pageSetup')));
        expect(xml, isNot(contains('<printOptions')));
      }
    });

    test('preserves printer settings r:id and autoPageBreaks', () {
      final base = Excel.createExcel();
      base['Sheet1'].cell(CellIndex.indexByString('A1')).value =
          TextCellValue('x');
      final patched = _patchSheetXml(base.encode()!, (xml) {
        return xml
            .replaceFirstMapped(
                RegExp(r'<worksheet[^>]*>'),
                (m) => '${m[0]}<sheetPr>'
                    '<outlinePr summaryBelow="0" summaryRight="0"/>'
                    '<pageSetUpPr autoPageBreaks="0" fitToPage="1"/></sheetPr>')
            .replaceFirst(
                RegExp(r'<pageMargins[^>]*/>'),
                '<pageMargins left="0.25" right="0.25" top="0.5" '
                    'bottom="0.5" header="0.2" footer="0.2"/>'
                    '<pageSetup paperSize="9" orientation="portrait" '
                    'fitToHeight="0" r:id="rId1"/>');
      });

      final excel = Excel.decodeBytes(patched);
      final sheet = excel['Sheet1'];
      expect(sheet.pageSetup!.paperSize, PaperSize.a4);
      expect(sheet.pageSetup!.fitToPage, isTrue);
      expect(sheet.pageSetup!.fitToHeight, 0);
      expect(sheet.pageMargins!.left, 0.25);

      sheet.setPageOrientation(PageOrientation.landscape);
      final xml = _sheetXml(excel.encode()!);
      expect(xml, contains('<pageSetUpPr autoPageBreaks="0" fitToPage="1"/>'));
      expect(
          xml,
          contains('<pageSetup paperSize="9" fitToHeight="0" '
              'orientation="landscape" r:id="rId1"/>'));
      expect(xml, contains('<pageMargins left="0.25"'));
    });

    test('clearing page setup removes fitToPage but keeps other sheetPr data',
        () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.fitToPages();
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
      final decoded = Excel.decodeBytes(excel.encode()!);
      decoded['Sheet1'].clearPageSetup();

      final xml = _sheetXml(decoded.encode()!);
      expect(xml, isNot(contains('pageSetUpPr')));
      expect(xml, isNot(contains('<pageSetup')));
      expect(xml, contains('<outlinePr'));
    });

    test('header and footer text is escaped once and read back unchanged', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].headerFooter = HeaderFooter(
        oddHeader: '&CExample "Q3" report',
        oddFooter: '&RPage &P of &N',
      );
      final bytes = excel.encode()!;

      final xml = _sheetXml(bytes);
      expect(xml, contains('<oddHeader>&amp;CExample "Q3" report</oddHeader>'));
      expect(xml, contains('<oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter>'));

      final headerFooter = Excel.decodeBytes(bytes)['Sheet1'].headerFooter!;
      expect(headerFooter.oddHeader, '&CExample "Q3" report');
      expect(headerFooter.oddFooter, '&RPage &P of &N');
    });
  });
}
