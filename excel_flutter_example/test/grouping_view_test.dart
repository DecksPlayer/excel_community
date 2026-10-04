import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/services/helpers/grouping_helper.dart';
import 'package:excel_flutter_example/widgets/grouping_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  Future<void> pumpView(WidgetTester tester, Size size) async {
    tester.view.physicalSize = size;
    tester.view.devicePixelRatio = 1.0;
    addTearDown(tester.view.reset);
    await tester.pumpWidget(MaterialApp(
      home: Scaffold(
        body: SingleChildScrollView(
          padding: const EdgeInsets.all(24),
          child: GroupingView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
        ),
      ),
    ));
  }

  for (final size in const [Size(1400, 3000), Size(400, 7000)]) {
    testWidgets('renders every tab without layout errors at $size', (tester) async {
      await pumpView(tester, size);
      for (final tab in ['Column Groups', 'Options & Code']) {
        await tester.tap(find.textContaining(tab));
        await tester.pumpAndSettle();
      }
      expect(find.textContaining('hidden rows → [1, 2, 3]'), findsOneWidget);
    });
  }

  testWidgets('the + button expands a collapsed group through the library', (tester) async {
    await pumpView(tester, const Size(1400, 3000));
    final card = find.ancestor(of: find.text('Start Collapsed'), matching: find.byType(Card));
    Finder inCard(String text) => find.descendant(of: card, matching: find.text(text));

    expect(inCard('Jan'), findsNothing, reason: 'rows hidden while collapsed');
    await tester.tap(inCard('+'));
    await tester.pumpAndSettle();
    expect(inCard('Jan'), findsOneWidget);
    expect(inCard('−'), findsOneWidget);

    await tester.tap(inCard('−'));
    await tester.pumpAndSettle();
    expect(inCard('Jan'), findsNothing);
  });

  test('demo workbook keeps its groups after save and reopen', () {
    final decoded = Excel.decodeBytes(buildGroupingWorkbook().encode()!)['Sales 2026'];
    expect(decoded.rowGroups.where((g) => g.level == 1).length, 2, reason: 'H1 and H2');
    expect(decoded.rowGroups.where((g) => g.level == 2).length, 4, reason: 'one per quarter');
    expect(decoded.rowGroups.firstWhere((g) => g.level == 1 && g.start > 1).collapsed, isTrue);
    expect(decoded.columnGroups.single, const OutlineGroup(1, 4, 1));
  });
}
