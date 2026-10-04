const String groupingSnippet = r'''
import 'package:excel_community/excel_community.dart';

void groupReport(Excel excel) {
  final sheet = excel['Report'];

  // Rows 2-9 in a group, with a nested group for rows 3-5
  sheet.groupRows(1, 8);
  sheet.groupRows(2, 4, collapsed: true);

  // Columns B-D grouped and collapsed
  sheet.groupColumns(1, 3, collapsed: true);

  // Expand / collapse later
  sheet.expandColumnGroup(1, 3);

  for (final group in sheet.rowGroups) {
    print('${group.start}..${group.end} level ${group.level}');
  }
}
''';
