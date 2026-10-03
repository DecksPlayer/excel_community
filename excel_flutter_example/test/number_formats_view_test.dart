import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/data/number_format_catalog.dart';
import 'package:excel_flutter_example/services/helpers/styles_helper.dart';
import 'package:excel_flutter_example/widgets/number_formats_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

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
            child: NumberFormatsView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
          ),
        ),
      ),
    );
  }

  for (final size in const [Size(1400, 4000), Size(400, 9000)]) {
    testWidgets('renders every category without layout errors at $size', (tester) async {
      await pumpView(tester, size);
      expect(find.text('Number Formats Wiki'), findsOneWidget);
      expect(find.text('1,234.57'), findsOneWidget); // standard_4 rendered live

      for (final category in NumFormatCategory.values) {
        await tester.tap(find.textContaining(category.label));
        await tester.pumpAndSettle();
        final count = numFormatCatalog.where((e) => e.category == category).length;
        expect(find.byIcon(Icons.copy), findsNWidgets(count), reason: category.label);
      }
    });
  }

  testWidgets('search looks across all categories', (tester) async {
    await pumpView(tester, const Size(1400, 4000));
    await tester.enterText(find.byType(TextField), 'accounting');
    await tester.pumpAndSettle();
    final matches = numFormatCatalog.where((e) => e.name.toLowerCase().contains('accounting')).length;
    expect(find.text('$matches matches across all categories'), findsOneWidget);

    await tester.enterText(find.byType(TextField), '14');
    await tester.pumpAndSettle();
    expect(find.text('Short Date'), findsOneWidget);
  });

  test('demo workbook keeps every number format after save and reload', () {
    final bytes = buildNumberFormatsWorkbook().encode()!;
    final decoded = Excel.decodeBytes(bytes);

    expect(decoded.tables.keys, [for (final c in NumFormatCategory.values) c.label]);

    for (final category in NumFormatCategory.values) {
      final sheet = decoded[category.label];
      var row = 3;
      for (final entry in numFormatCatalog.where((e) => e.category == category)) {
        for (final _ in entry.samples) {
          final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: row));
          expect(cell.cellStyle?.numberFormat.formatCode, entry.format.formatCode,
              reason: '${category.label} row ${row + 1}: ${entry.name}');
          row++;
        }
      }
    }
  });
}
