const String pageSetupSnippet = '''
import 'package:excel_community/excel_community.dart';

void generatePageSetupExcel() {
  final excel = Excel.createExcel();

  // 1. One worksheet per paper size, each printing on a different paper
  for (final paper in PaperSize.values) {
    final sheet = excel[paper.name]; // "Letter", "A4", "Legal", ...
    sheet.setPaperSize(paper);
    sheet.setPageOrientation(PageOrientation.portrait);
    sheet.setPrintGridLines(true);
    sheet.setPrintCentered(horizontally: true);
    sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue(
        '\${paper.name} - \${paper.widthMm} x \${paper.heightMm} mm (code \${paper.code})');
  }

  // Any other SpreadsheetML paper code is also supported
  excel['Custom'].setPaperSize(PaperSize.fromCode(41)); // German Legal Fanfold

  // 2. Wide sales report: A4 landscape, all columns on one page width
  final report = excel['Sales Report'];
  report.setPageOrientation(PageOrientation.landscape);
  report.setPaperSize(PaperSize.a4);
  report.fitToPages(width: 1, height: 0); // 0 = as many pages tall as needed
  report.pageMargins = PageMargins.narrow;
  report.setPrintGridLines(true);
  report.headerFooter = HeaderFooter(oddFooter: '&CPage &P of &N');

  // 3. Invoice: US Letter portrait at 90% scale, custom margins in cm
  final invoice = excel['Invoice'];
  invoice.pageSetup = const PageSetup(
    orientation: PageOrientation.portrait,
    paperSize: PaperSize.letter,
    scale: 90,
    blackAndWhite: true,
    errors: PrintErrors.blank, // print #N/A, #DIV/0! as blank
  );
  invoice.pageMargins = PageMargins.fromCentimeters(
    left: 2, right: 2, top: 2.5, bottom: 2.5,
  );

  // 4. Audit log: Legal paper, numbering starts at 10, comments at the end
  final audit = excel['Audit Log'];
  audit.pageSetup = const PageSetup(
    paperSize: PaperSize.legal,
    firstPageNumber: 10,
    useFirstPageNumber: true,
    pageOrder: PageOrder.overThenDown,
    cellComments: PrintCellComments.atEnd,
    copies: 2,
  );
  audit.setPrintHeadings(true); // print A, B, C / 1, 2, 3

  // 5. Inspect settings (also works on files you decode)
  final setup = report.pageSetup;
  print('\${setup?.paperSize?.name} \${setup?.orientation?.name}'); // "A4 landscape"

  // To reset:
  // report.clearPageSetup(); report.clearPageMargins(); report.clearPrintOptions();

  excel.save(fileName: 'print_ready_reports.xlsx');
}
''';
