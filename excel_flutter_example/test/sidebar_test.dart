import 'package:excel_flutter_example/models/section_detail.dart';
import 'package:excel_flutter_example/widgets/sidebar.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  testWidgets('sidebar items render inside a decorated container without ink warnings',
      (tester) async {
    tester.view.physicalSize = const Size(1200, 3000);
    tester.view.devicePixelRatio = 1.0;
    addTearDown(tester.view.reset);

    var selected = SelectedSection.pageSetup;
    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          // Same setup as main.dart: a white, decorated sidebar container.
          body: Container(
            width: 260,
            decoration: const BoxDecoration(color: Colors.white),
            child: StatefulBuilder(
              builder: (context, setState) => Sidebar(
                selectedSection: selected,
                onSectionSelected: (section) => setState(() => selected = section),
              ),
            ),
          ),
        ),
      ),
    );

    expect(tester.takeException(), isNull);
    expect(find.byType(ListTile), findsWidgets);

    await tester.tap(find.text('Number Formatting'));
    await tester.pumpAndSettle();
    expect(selected, SelectedSection.numberFormats);
    expect(tester.takeException(), isNull);
  });
}
