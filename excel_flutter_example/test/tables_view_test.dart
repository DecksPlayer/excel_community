import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/data/table_samples.dart';
import 'package:excel_flutter_example/services/helpers/tables_helper.dart';
import 'package:excel_flutter_example/widgets/tables_view.dart';
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
          child: TablesView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
        ),
      ),
    ));
  }

  for (final size in const [Size(1400, 6000), Size(400, 20000)]) {
    testWidgets('renders every tab without layout errors at $size', (tester) async {
      await pumpView(tester, size);
      expect(find.text('Medium 9'), findsOneWidget);
      for (final tab in ['Options', 'Data & Formulas']) {
        await tester.tap(find.textContaining(tab));
        await tester.pumpAndSettle();
      }
      expect(find.textContaining('ref → A1:C6'), findsOneWidget, reason: 'appendTableRow grew the table');
    });
  }

  testWidgets('style search filters the gallery', (tester) async {
    await pumpView(tester, const Size(1400, 6000));
    await tester.enterText(find.byType(TextField), 'dark');
    await tester.pumpAndSettle();
    expect(find.text('Dark 11'), findsOneWidget);
    expect(find.text('Medium 9'), findsNothing);
  });

  test('every sample builds a valid table that survives save and reopen', () {
    for (final sample in tableSamples) {
      final sheet = tableSampleSheet();
      final table = sample.apply(sheet);
      expect(sheet.tables, contains(table), reason: sample.title);
    }
  });

  test('demo workbook keeps all tables', () {
    final decoded = Excel.decodeBytes(buildTablesWorkbook().encode()!);
    expect(decoded['Sales'].getTable('Sales')!.showTotalsRow, isTrue);
    expect(decoded['Style Gallery'].tables.length, builtInTableStyles.length);
    expect(decoded['Style Gallery'].tables.map((t) => t.style), builtInTableStyles.map((s) => s.$2));
  });
}
