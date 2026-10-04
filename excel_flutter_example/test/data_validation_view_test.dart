import 'package:excel_community/excel_community.dart';
import 'package:excel_flutter_example/data/data_validation_samples.dart';
import 'package:excel_flutter_example/services/helpers/data_validation_helper.dart';
import 'package:excel_flutter_example/widgets/data_validation_view.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  for (final size in const [Size(1400, 3000), Size(400, 7000)]) {
    testWidgets('renders every tab without layout errors at $size', (tester) async {
      tester.view.physicalSize = size;
      tester.view.devicePixelRatio = 1.0;
      addTearDown(tester.view.reset);
      await tester.pumpWidget(MaterialApp(
        home: Scaffold(
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: DataValidationView(isGenerating: false, onGenerate: () {}, status: 'Ready'),
          ),
        ),
      ));

      expect(find.text('Pick a status'), findsOneWidget, reason: 'input message box');
      expect(find.text('rejected'), findsWidgets);

      for (final tab in ['Numbers, Dates & Times', 'Text, Custom & Checks']) {
        await tester.tap(find.textContaining(tab));
        await tester.pumpAndSettle();
      }
      expect(find.text('evaluated by Excel'), findsWidgets, reason: 'custom formula');
    });
  }

  test('demo workbook keeps every rule after save and reopen', () {
    final decoded = Excel.decodeBytes(buildDataValidationWorkbook().encode()!);
    final tasks = decoded['Tasks'];
    expect(tasks.dataValidations.length, 11);
    for (final sample in validationSamples.where((s) => s.title != 'Check in Dart')) {
      expect(tasks.dataValidations.values, contains(sample.rule), reason: sample.title);
    }
    // The sample row satisfies every rule that can be checked in Dart.
    for (var col = 0; col < 11; col++) {
      final cell = tasks.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 1));
      expect(cell.dataValidation?.accepts(cell.value), isNot(false), reason: 'column $col');
    }
  });
}
