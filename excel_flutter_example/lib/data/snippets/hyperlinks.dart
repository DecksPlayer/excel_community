const String hyperlinksSnippet = r'''
import 'package:excel_community/excel_community.dart';

void addLinks(Excel excel) {
  final sheet = excel['Links'];

  sheet.setHyperlink(CellIndex.indexByString('A1'),
      Hyperlink.url('https://pub.dev', tooltip: 'Open pub.dev'));
  sheet.setHyperlink(CellIndex.indexByString('A2'),
      Hyperlink.email('sales@example.com', subject: 'Q3 report'),
      text: 'Contact sales');
  sheet.setHyperlink(CellIndex.indexByString('A3'),
      Hyperlink.cell('Q1 Sales', 'B4'), text: 'Go to Q1');
  sheet.setHyperlinkRange('A5:C5', Hyperlink.location('TotalSales'));

  final link = sheet.cell(CellIndex.indexByString('A1')).hyperlink;
  print(link?.url); // https://pub.dev
}
''';
