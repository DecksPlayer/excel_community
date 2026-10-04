import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generatePageSetupHelper() async {
  final excel = Excel.createExcel();

  // Overview: every named paper size with its code and dimensions.
  final overview = excel['Overview'];
  overview.setPrintGridLines(true);
  overview.setPrintCentered(horizontally: true);
  _buildTable(
    overview,
    title: 'Paper Sizes - one worksheet per size (open File > Print to compare)',
    themeColor: '#0EA5E9',
    headers: ['Paper', 'Code', 'Width (mm)', 'Height (mm)', 'Width (in)', 'Height (in)'],
    rows: [
      for (final paper in PaperSize.values)
        [
          paper.name,
          '${paper.code}',
          _fmt(paper.widthMm!),
          _fmt(paper.heightMm!),
          _fmt(paper.widthMm! / 25.4),
          _fmt(paper.heightMm! / 25.4),
        ],
    ],
  );

  // One worksheet per paper size, so each prints on a different paper.
  for (final paper in PaperSize.values) {
    final sheet = excel[paper.name];
    sheet.setPaperSize(paper);
    sheet.setPageOrientation(PageOrientation.portrait);
    sheet.setPrintGridLines(true);
    sheet.setPrintCentered(horizontally: true);
    sheet.headerFooter = HeaderFooter(oddFooter: '&C${paper.name} - Page &P of &N');
    _buildTable(
      sheet,
      title: '${paper.name} - ${_fmt(paper.widthMm!)} x ${_fmt(paper.heightMm!)} mm (code ${paper.code})',
      themeColor: '#0EA5E9',
      headers: ['Item', 'Q1', 'Q2', 'Q3', 'Q4'],
      rows: List.generate(
        60,
        (i) => ['Item ${i + 1}', for (var q = 0; q < 4; q++) '${100 + (i * 31 + q * 17) % 400}'],
      ),
    );
  }

  // 1. Wide sales report: A4 landscape, fitted to one page wide.
  final report = excel['Sales Report'];
  report.setPageOrientation(PageOrientation.landscape);
  report.setPaperSize(PaperSize.a4);
  report.fitToPages(width: 1, height: 0);
  report.pageMargins = PageMargins.narrow;
  report.setPrintGridLines(true);
  report.setPrintCentered(horizontally: true);
  report.headerFooter = HeaderFooter(
    oddHeader: '&LSales Report&RPrinted &D',
    oddFooter: '&CPage &P of &N',
  );
  _buildTable(
    report,
    title: 'Regional Sales - A4 Landscape, Fit to 1 Page Wide',
    themeColor: '#0EA5E9',
    headers: [
      'Region', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Total',
    ],
    rows: List.generate(40, (i) {
      final months = List.generate(12, (m) => 1000 + (i * 37 + m * 53) % 900);
      return [
        'Region ${i + 1}',
        ...months.map((v) => '$v'),
        '${months.reduce((a, b) => a + b)}',
      ];
    }),
  );

  // 2. Invoice: Letter portrait at 90 %, black & white, margins in cm.
  final invoice = excel['Invoice'];
  invoice.pageSetup = const PageSetup(
    orientation: PageOrientation.portrait,
    paperSize: PaperSize.letter,
    scale: 90,
    blackAndWhite: true,
    errors: PrintErrors.blank,
  );
  invoice.pageMargins =
      PageMargins.fromCentimeters(left: 2, right: 2, top: 2.5, bottom: 2.5);
  _buildTable(
    invoice,
    title: 'Invoice #2026-1003 - Letter Portrait, 90 % Scale',
    themeColor: '#0F172A',
    headers: ['Item', 'Description', 'Qty', 'Unit Price', 'Amount'],
    rows: [
      ['SKU-001', 'Annual support plan', '1', '\$4,800.00', '\$4,800.00'],
      ['SKU-014', 'Onboarding workshop', '2', '\$1,250.00', '\$2,500.00'],
      ['SKU-022', 'Additional seats', '25', '\$36.00', '\$900.00'],
      ['', '', '', 'Total', '\$8,200.00'],
    ],
  );

  // 3. Audit log: Legal paper, numbering starts at 10, 2 copies, headings.
  final audit = excel['Audit Log'];
  audit.pageSetup = const PageSetup(
    paperSize: PaperSize.legal,
    firstPageNumber: 10,
    useFirstPageNumber: true,
    pageOrder: PageOrder.overThenDown,
    cellComments: PrintCellComments.atEnd,
    copies: 2,
  );
  audit.setPrintHeadings(true);
  _buildTable(
    audit,
    title: 'Audit Log - Legal, Starts at Page 10, Row/Column Headings',
    themeColor: '#8B5CF6',
    headers: ['Timestamp', 'User', 'Action', 'Resource'],
    rows: List.generate(
      25,
      (i) => [
        '2026-10-03 0${i % 10}:${(i * 7 % 60).toString().padLeft(2, '0')}',
        'user${i % 5}@example.com',
        i.isEven ? 'UPDATE' : 'READ',
        '/reports/${1000 + i}',
      ],
    ),
  );

  final summaryInfo = 'Overview + ${PaperSize.values.length} paper size sheets, '
      'Sales Report (A4 landscape, fit 1 page wide), '
      'Invoice (Letter, 90 %), Audit Log (Legal, from page 10)';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'print_ready_reports.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Page Setup spreadsheet generated successfully!\n'
          'Worksheets: $summaryInfo\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 File: print_ready_reports.xlsx\n'
          '💡 Open File → Print in Excel to see each sheet\'s page layout!';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Page Setup Showcase',
      fileName: 'print_ready_reports.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Page Setup spreadsheet saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summaryInfo\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 Open File → Print in Excel to preview the page layouts!';
    }
    return 'Save cancelled.';
  }
}

void _buildTable(
  Sheet sheet, {
  required String title,
  required String themeColor,
  required List<String> headers,
  required List<List<String>> rows,
}) {
  sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue(title);
  sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
    bold: true,
    fontSize: 14,
    fontColorHex: ExcelColor.fromHexString(themeColor),
  );

  for (int col = 0; col < headers.length; col++) {
    final cell =
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 2));
    cell.value = TextCellValue(headers[col]);
    cell.cellStyle = CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      backgroundColorHex: ExcelColor.fromHexString(themeColor),
      horizontalAlign: HorizontalAlign.Center,
    );
    sheet.setColumnWidth(col, col == 0 ? 20.0 : 12.0);
  }

  for (int r = 0; r < rows.length; r++) {
    for (int c = 0; c < rows[r].length; c++) {
      sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 3))
          .value = TextCellValue(rows[r][c]);
    }
  }
}

String _fmt(double v) {
  final s = v.toStringAsFixed(1);
  return s.endsWith('.0') ? s.substring(0, s.length - 2) : s;
}
