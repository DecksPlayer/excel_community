import 'package:excel_community/excel_community.dart';

/// One hyperlink example: its snippet and the code that applies it.
class HyperlinkSample {
  final String category; // external, internal, manage
  final String title;
  final String subtitle;
  final String code;
  final void Function(Excel excel) apply;

  const HyperlinkSample({
    required this.category,
    required this.title,
    required this.subtitle,
    required this.code,
    required this.apply,
  });
}

/// Result of running a [HyperlinkSample] through the library.
class HyperlinkSampleResult {
  final String cellText;
  final bool styled;
  final Hyperlink? link;

  /// Hyperlinks read back after saving and reopening the workbook.
  final Map<String, Hyperlink> reopened;

  /// Extra output for the "manage" examples.
  final String? note;

  const HyperlinkSampleResult(this.cellText, this.styled, this.link, this.reopened, this.note);
}

CellIndex _a1 = CellIndex.indexByString('A1');

final List<HyperlinkSample> hyperlinkSamples = [
  // --- External ---------------------------------------------------------------
  HyperlinkSample(
    category: 'external',
    title: 'Web Page',
    subtitle: 'Hyperlink.url()',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.url('https://pub.dev/packages/excel_community',\n"
        "      tooltip: 'Open on pub.dev'),\n"
        "  text: 'excel_community on pub.dev',\n"
        ');',
    apply: (excel) => excel['Links'].setHyperlink(
      _a1,
      Hyperlink.url('https://pub.dev/packages/excel_community', tooltip: 'Open on pub.dev'),
      text: 'excel_community on pub.dev',
    ),
  ),
  HyperlinkSample(
    category: 'external',
    title: 'E-mail',
    subtitle: 'Hyperlink.email(subject:)',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.email('sales@example.com',\n"
        "      subject: 'Q3 report'),\n"
        ');\n'
        '// Empty cell: the address is used as text',
    apply: (excel) => excel['Links']
        .setHyperlink(_a1, Hyperlink.email('sales@example.com', subject: 'Q3 report')),
  ),
  HyperlinkSample(
    category: 'external',
    title: 'Local File',
    subtitle: 'Hyperlink.url(path)',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.url('reports/2026/q3.pdf'),\n"
        "  text: 'Q3 report (PDF)',\n"
        ');\n'
        '// Relative to the workbook location',
    apply: (excel) => excel['Links']
        .setHyperlink(_a1, Hyperlink.url('reports/2026/q3.pdf'), text: 'Q3 report (PDF)'),
  ),
  HyperlinkSample(
    category: 'external',
    title: 'Page with Anchor',
    subtitle: 'url + location',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink(url: 'https://dart.dev/language',\n"
        "      location: 'variables'),\n"
        "  text: 'Dart variables',\n"
        ');',
    apply: (excel) => excel['Links'].setHyperlink(
        _a1, const Hyperlink(url: 'https://dart.dev/language', location: 'variables'),
        text: 'Dart variables'),
  ),

  // --- Internal ---------------------------------------------------------------
  HyperlinkSample(
    category: 'internal',
    title: 'Cell on Another Sheet',
    subtitle: 'Hyperlink.cell()',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.cell('Q1 Sales', 'B4'),\n"
        "  text: 'Go to Q1 total',\n"
        ');\n'
        "// Sheet names with spaces are quoted: 'Q1 Sales'!B4",
    apply: (excel) =>
        excel['Links'].setHyperlink(_a1, Hyperlink.cell('Q1 Sales', 'B4'), text: 'Go to Q1 total'),
  ),
  HyperlinkSample(
    category: 'internal',
    title: 'Range on Another Sheet',
    subtitle: "Hyperlink.cell(sheet, 'A1:C3')",
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.cell('Data', 'A1:C3',\n"
        "      tooltip: 'Selects the data block'),\n"
        "  text: 'Show data',\n"
        ');',
    apply: (excel) => excel['Links'].setHyperlink(
        _a1, Hyperlink.cell('Data', 'A1:C3', tooltip: 'Selects the data block'),
        text: 'Show data'),
  ),
  HyperlinkSample(
    category: 'internal',
    title: 'Defined Name',
    subtitle: 'Hyperlink.location()',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.location('TotalSales'),\n"
        "  text: 'Total sales',\n"
        ');\n'
        '// The name must exist in the workbook',
    apply: (excel) =>
        excel['Links'].setHyperlink(_a1, Hyperlink.location('TotalSales'), text: 'Total sales'),
  ),
  HyperlinkSample(
    category: 'internal',
    title: 'Link over a Range',
    subtitle: 'setHyperlinkRange()',
    code: "sheet.setHyperlinkRange(\n"
        "  'A1:C1',\n"
        "  Hyperlink.cell('Summary', 'A1'),\n"
        ');\n'
        '// One <hyperlink ref="A1:C1"> for the 3 cells',
    apply: (excel) {
      final sheet = excel['Links'];
      sheet.cell(_a1).value = TextCellValue('Back to summary');
      sheet.setHyperlinkRange('A1:C1', Hyperlink.cell('Summary', 'A1'));
    },
  ),

  // --- Manage -----------------------------------------------------------------
  HyperlinkSample(
    category: 'manage',
    title: 'From the Cell',
    subtitle: 'cell.hyperlink',
    code: "final cell = sheet.cell(CellIndex.indexByString('A1'));\n"
        "cell.hyperlink = Hyperlink.url('https://dart.dev');\n"
        'print(cell.hyperlink?.url);',
    apply: (excel) => excel['Links'].cell(_a1).hyperlink = Hyperlink.url('https://dart.dev'),
  ),
  HyperlinkSample(
    category: 'manage',
    title: 'Plain Style',
    subtitle: 'styled: false',
    code: "sheet.setHyperlink(\n"
        "  CellIndex.indexByString('A1'),\n"
        "  Hyperlink.url('https://dart.dev'),\n"
        "  text: 'Keeps the cell style',\n"
        '  styled: false,\n'
        ');',
    apply: (excel) => excel['Links'].setHyperlink(_a1, Hyperlink.url('https://dart.dev'),
        text: 'Keeps the cell style', styled: false),
  ),
  HyperlinkSample(
    category: 'manage',
    title: 'Links Move with Rows',
    subtitle: 'insertRow() / removeRow()',
    code: "sheet.setHyperlink(CellIndex.indexByString('A1'),\n"
        "    Hyperlink.url('https://dart.dev'));\n"
        'sheet.insertRow(0);\n'
        'print(sheet.hyperlinks.keys); // (A2)',
    apply: (excel) {
      final sheet = excel['Links'];
      sheet.setHyperlink(_a1, Hyperlink.url('https://dart.dev'));
      sheet.insertRow(0);
    },
  ),
  HyperlinkSample(
    category: 'manage',
    title: 'Remove Links',
    subtitle: 'removeHyperlink() / clearHyperlinks()',
    code: "sheet.removeHyperlink(CellIndex.indexByString('A1'));\n"
        'sheet.clearHyperlinks();\n'
        '// The cell value and style are kept',
    apply: (excel) {
      final sheet = excel['Links'];
      sheet.setHyperlink(_a1, Hyperlink.url('https://dart.dev'), text: 'Was a link');
      sheet.removeHyperlink(_a1);
    },
  ),
];

final Map<HyperlinkSample, HyperlinkSampleResult> _hyperlinkSampleCache = {};

/// Runs [sample] on a fresh workbook, then saves and reopens it to read
/// the links back from the file. Results are cached so zip encode/decode
/// only runs once across the app session.
HyperlinkSampleResult runHyperlinkSample(HyperlinkSample sample) {
  return _hyperlinkSampleCache.putIfAbsent(sample, () {
    final excel = Excel.createExcel();
    excel.rename('Sheet1', 'Links');
    sample.apply(excel);
    final sheet = excel['Links'];

    // First cell holding a value (A1, or A2 after insertRow).
    final cellIndex = sheet.hyperlinks.keys.isNotEmpty
        ? CellIndex.indexByString(sheet.hyperlinks.keys.first.split(':').first)
        : _a1;
    final cell = sheet.cell(cellIndex);

    final reopened = Excel.decodeBytes(excel.encode()!)['Links'].hyperlinks;

    return HyperlinkSampleResult(
      cell.value?.toString() ?? '',
      cell.cellStyle?.underline == Underline.Single,
      cell.hyperlink,
      reopened,
      switch (sample.title) {
        'Links Move with Rows' => 'sheet.hyperlinks.keys → ${sheet.hyperlinks.keys.toList()}',
        'Remove Links' => 'hasHyperlinks → ${sheet.hasHyperlinks}',
        'From the Cell' => 'cell.hyperlink?.url → ${cell.hyperlink?.url}',
        _ => null,
      },
    );
  });
}
