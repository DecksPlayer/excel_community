import 'package:xml/xml_events.dart' as xml_events;

void main() {
  const xml = '''
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
      <row r="1">
        <c r="A1" t="s">
          <v>0</v>
        </c>
        <c r="B1" t="n">
          <v>100</v>
        </c>
      </row>
    </sheetData>
  </worksheet>
  ''';

  final events = xml_events.parseEvents(xml);
  for (final event in events) {
    if (event is xml_events.XmlStartElementEvent) {
      final attrs =
          event.attributes.map((a) => '${a.name}=${a.value}').toList();
      print('Start Element: ${event.name} with attributes $attrs');
    } else if (event is xml_events.XmlEndElementEvent) {
      print('End Element: ${event.name}');
    } else if (event is xml_events.XmlTextEvent) {
      final text = event.value.trim();
      if (text.isNotEmpty) {
        print('Text: $text');
      }
    }
  }
}
