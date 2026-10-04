import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/data/hyperlink_samples.dart';
import 'package:excel_flutter_example/services/helpers/hyperlinks_helper.dart';
import 'package:excel_flutter_example/widgets/hyperlinks_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  for (final size in const [Size(1400, 3000), Size(400, 6000)]) {
    testWidgets('renders every tab without layout errors at $size', (tester) async {
      tester.view.physicalSize = size;
      tester.view.devicePixelRatio = 1.0;
      addTearDown(tester.view.reset);
      await tester.pumpWidget(MaterialApp(
        home: Scaffold(
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: HyperlinksView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
          ),
        ),
      ));

      expect(find.text('excel_community on pub.dev'), findsOneWidget);
      expect(find.text('sales@example.com'), findsOneWidget, reason: 'default text of an e-mail link');

      await tester.tap(find.textContaining('Inside the Workbook'));
      await tester.pumpAndSettle();
      expect(find.textContaining("location: 'Q1 Sales'!B4"), findsOneWidget);

      await tester.tap(find.textContaining('Read, Style & Remove'));
      await tester.pumpAndSettle();
      expect(find.textContaining('sheet.hyperlinks.keys → [A2]'), findsOneWidget);
    });
  }

  test('every sample survives a save and reopen', () {
    for (final sample in hyperlinkSamples) {
      final result = runHyperlinkSample(sample);
      if (sample.title == 'Remove Links') {
        expect(result.reopened, isEmpty);
      } else {
        expect(result.reopened.values, contains(result.link), reason: sample.title);
      }
    }
  });

  test('demo workbook links both ways between its sheets', () {
    final decoded = Excel.decodeBytes(buildHyperlinksWorkbook().encode()!);
    expect(decoded['Links'].hyperlinks.length, 5);
    expect(decoded['Links'].getHyperlink(CellIndex.indexByString('B6'))?.location, "'Q1 Sales'!B4");
    expect(decoded['Q1 Sales'].getHyperlink(CellIndex.indexByString('D1'))?.location, 'Links!A1');
  });
}
