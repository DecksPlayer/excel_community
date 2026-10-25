import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel_community/src/parser/fast_xml_events.dart';
import 'package:test/test.dart';
import 'package:xml/xml_events.dart' as xml_events;

/// Events of [input] as comparable strings: node type, `toString()` and, for
/// elements, the attribute quote types; or the error type when parsing fails.
List<String> _describe(Iterable<xml_events.XmlEvent> Function(String) parse, String input) {
  final out = <String>[];
  try {
    for (final e in parse(input)) {
      final quotes = e is xml_events.XmlStartElementEvent
          ? e.attributes.map((a) => a.attributeType.name).join(',')
          : '';
      out.add('${e.nodeType.name} ${e.toString()} $quotes');
    }
  } catch (e) {
    out.add('error ${e.runtimeType}');
  }
  return out;
}

void _expectSameEvents(String input, {String? reason}) {
  expect(_describe(fastXmlEvents, input), _describe(xml_events.parseEvents, input), reason: reason);
}

void main() {
  test('same events as xml_events.parseEvents for every part of the test files', () {
    var parts = 0;
    for (final file in Directory('test/test_resources').listSync().whereType<File>()) {
      if (!file.path.endsWith('.xlsx') || file.path.contains('encrypted')) continue;
      final Archive archive;
      try {
        archive = ZipDecoder().decodeBytes(file.readAsBytesSync());
      } catch (_) {
        continue;
      }
      for (final part in archive.files) {
        if (!part.isFile || !(part.name.endsWith('.xml') || part.name.endsWith('.rels'))) continue;
        _expectSameEvents(utf8.decode(part.content, allowMalformed: true),
            reason: '${file.uri.pathSegments.last} ${part.name}');
        parts++;
      }
    }
    expect(parts, greaterThan(100));
  });

  test('same events for edge cases, including the ones it hands over', () {
    const cases = [
      '',
      'plain text',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a/>',
      "<?xml version='1.0'?><a/>",
      '<a b="1" c=\'2\' d = "3"  >x</a >',
      '<x:c r="A1" s="2" t="s"><x:v>0</x:v></x:c>',
      '<t xml:space="preserve">  Fish &amp; Chips &lt;3 &#65;&#x42; &quot;q&quot; &apos;a&apos;</t>',
      '<a title="&lt;b&gt; &amp; &#10;"/>',
      '<a>Ñandú — 日本語 🙂</a>',
      '<a><!-- a comment --><![CDATA[<raw> & stuff]]></a>',
      '<a>\n  <b/>\n  <c  />\n</a>',
      '<a b="x>y" c="it\'s"/>',
      '<?mso-application progid="Excel.Sheet"?><a/>',
      '<!DOCTYPE a><a/>',
      '<a b=unquoted/>',
      '<a b/>',
      '<a b="1"c="2"/>',
      '<a',
      '<a b="1',
      '</a',
      '<1a/>',
      '<a>text<',
      'x & y < z',
      '<a>&unknown;</a>',
    ];
    for (final input in cases) {
      _expectSameEvents(input, reason: input);
    }
    // The comparison sees real events, not two empty lists.
    expect(_describe(fastXmlEvents, "<a b='&amp;'>x &lt; y</a>"), [
      "ELEMENT <a b='&amp;'> SINGLE_QUOTE",
      'TEXT x &lt; y ',
      'ELEMENT </a> ',
    ]);
    expect(_describe(fastXmlEvents, '<a'), ['error XmlParserException']);
  });
}
