import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/services/helpers/data_export_helper.dart';
import 'package:excel_flutter_example/widgets/data_export_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  for (final size in const [Size(1400, 3000), Size(400, 6000)]) {
    testWidgets('renders both tabs without layout errors at $size', (tester) async {
      tester.view.physicalSize = size;
      tester.view.devicePixelRatio = 1.0;
      addTearDown(tester.view.reset);
      await tester.pumpWidget(MaterialApp(
        home: Scaffold(
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: DataExportView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
          ),
        ),
      ));

      // Live library output: displayed values and RFC 4180 quoting.
      expect(find.textContaining('95.0%'), findsWidgets);
      expect(find.textContaining('"Smith, Luis"'), findsWidgets);

      await tester.tap(find.text('Import (3)'));
      await tester.pumpAndSettle();
      expect(find.text('Round Trip'), findsOneWidget);
      expect(find.textContaining('Copy!D2 = DateCellValue'), findsOneWidget);
    });
  }

  test('demo workbook contains the exports and the imported sheet', () {
    final decoded = Excel.decodeBytes(buildDataExportWorkbook().encode()!);
    expect(decoded.tables.keys,
        containsAll(['Customers', 'Orders', 'JSON Export', 'CSV Export', 'Imported from JSON']));
    expect(decoded['Imported from JSON'].rowsAsMaps(), [
      {'Name': 'Eva', 'Age': 40, 'City': 'Lima', 'Joined': null},
      {'Name': 'Tom', 'Age': 35, 'City': null, 'Joined': '2026-01-15'},
    ]);
    expect(decoded['CSV Export'].cell(CellIndex.indexByString('A3')).value.toString(),
        'Name,Age,Active,Joined,Score');
  });
}
