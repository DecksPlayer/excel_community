import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:excel_flutter_example/widgets/page_setup_view.dart';

void main() {
  Future<void> pumpView(WidgetTester tester, Size size) async {
    tester.view.physicalSize = size;
    tester.view.devicePixelRatio = 1.0;
    addTearDown(tester.view.reset);
    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: PageSetupView(
              isGenerating: false,
              onGenerate: () {},
              status: 'Ready',
            ),
          ),
        ),
      ),
    );
  }

  for (final size in const [Size(1400, 2400), Size(400, 4000)]) {
    testWidgets('renders every tab without layout errors at $size', (tester) async {
      await pumpView(tester, size);
      expect(find.text('Page Setup & Print Wiki'), findsOneWidget);
      expect(find.text('A4'), findsOneWidget);
      expect(find.textContaining('sheet.setPaperSize(PaperSize.a4);'), findsOneWidget);

      await tester.tap(find.text('Landscape'));
      await tester.pumpAndSettle();

      for (final tab in ['Orientation & Scaling', 'Margins', 'Print Options']) {
        await tester.tap(find.text(tab));
        await tester.pumpAndSettle();
      }
      expect(find.text('Black & White'), findsOneWidget);
      expect(find.textContaining('.copyWith(blackAndWhite: true);'), findsOneWidget);
    });
  }

  testWidgets('search filters paper sizes', (tester) async {
    await pumpView(tester, const Size(1400, 2400));
    await tester.enterText(find.byType(TextField), 'envelope');
    await tester.pumpAndSettle();
    expect(find.text('Envelope DL'), findsOneWidget);
    expect(find.text('A4'), findsNothing);
  });
}
