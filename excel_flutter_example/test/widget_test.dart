import 'package:flutter_test/flutter_test.dart';
import 'package:excel_flutter_example/main.dart';

void main() {
  testWidgets('Excel gallery app renders successfully', (WidgetTester tester) async {
    // Build our app and trigger a frame.
    await tester.pumpWidget(const MyApp());

    // Verify that the title of the app is visible.
    expect(find.text('excel_community examples'), findsOneWidget);
  });
}
