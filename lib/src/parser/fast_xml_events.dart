import 'package:xml/xml.dart';
import 'package:xml/xml_events.dart' as xml_events;

/// The events of `xml_events.parseEvents(input)`, read with a hand-written
/// scanner.
///
/// The `xml` package parses with parser combinators, which made reading
/// worksheets and shared strings slow (most of the time when reading a file,
/// several times more when compiled to JavaScript). This scanner produces the
/// same events for the XML spreadsheets use: elements, attributes, text,
/// comments, CDATA and the XML declaration. Anything else (a DOCTYPE,
/// processing instructions, attributes without a quoted value, malformed
/// XML) hands the rest of the input to `xml_events.parseEvents`, so those
/// cases, errors included, behave exactly as before.
Iterable<xml_events.XmlEvent> fastXmlEvents(String input) =>
    _FastXmlEvents(input);

class _FastXmlEvents extends Iterable<xml_events.XmlEvent> {
  _FastXmlEvents(this._input);

  final String _input;

  @override
  Iterator<xml_events.XmlEvent> get iterator => _FastXmlEventIterator(_input);
}

class _FastXmlEventIterator implements Iterator<xml_events.XmlEvent> {
  _FastXmlEventIterator(this._input);

  static const _mapping = XmlDefaultEntityMapping.xml();

  static const _lt = 0x3C; // <
  static const _gt = 0x3E; // >
  static const _slash = 0x2F; // /
  static const _equals = 0x3D; // =
  static const _quote = 0x22; // "
  static const _apostrophe = 0x27; // '
  static const _question = 0x3F; // ?
  static const _bang = 0x21; // !

  final String _input;
  int _position = 0;
  xml_events.XmlEvent? _current;

  /// Parses the rest of the input once the scanner meets something it does
  /// not handle.
  Iterator<xml_events.XmlEvent>? _delegate;

  @override
  xml_events.XmlEvent get current => _current!;

  @override
  bool moveNext() {
    final delegate = _delegate;
    if (delegate != null) {
      final moved = delegate.moveNext();
      _current = moved ? delegate.current : null;
      return moved;
    }
    final input = _input;
    final start = _position;
    if (start >= input.length) {
      _current = null;
      return false;
    }
    final event =
        input.codeUnitAt(start) == _lt ? _markup(start) : _text(start);
    if (event == null) {
      _delegate = xml_events.parseEvents(input.substring(start)).iterator;
      return moveNext();
    }
    _current = event;
    return true;
  }

  xml_events.XmlEvent _text(int start) {
    var end = _input.indexOf('<', start);
    if (end == -1) end = _input.length;
    _position = end;
    return xml_events.XmlTextEvent(_decode(_input.substring(start, end)));
  }

  /// The event starting with `<` at [start], or `null` to delegate.
  xml_events.XmlEvent? _markup(int start) {
    final input = _input;
    if (start + 1 >= input.length) return null;
    final next = input.codeUnitAt(start + 1);
    if (next == _slash) return _endElement(start);
    if (next == _bang) {
      if (input.startsWith('<!--', start)) {
        return _delimited(start, '<!--', '-->', xml_events.XmlCommentEvent.new);
      }
      if (input.startsWith('<![CDATA[', start)) {
        return _delimited(start, '<![CDATA[', ']]>', xml_events.XmlCDATAEvent.new);
      }
      return null; // DOCTYPE
    }
    if (next == _question) return _declaration(start);
    return _startElement(start);
  }

  xml_events.XmlEvent? _delimited(int start, String open, String close,
      xml_events.XmlEvent Function(String) create) {
    final end = _input.indexOf(close, start + open.length);
    if (end == -1) return null;
    _position = end + close.length;
    return create(_input.substring(start + open.length, end));
  }

  xml_events.XmlStartElementEvent? _startElement(int start) {
    final input = _input;
    final nameEnd = _nameEnd(start + 1);
    if (nameEnd == -1) return null;
    final name = input.substring(start + 1, nameEnd);
    final attributes = <xml_events.XmlEventAttribute>[];
    final end = _attributes(nameEnd, attributes);
    if (end == -1) return null;
    // After the attributes: optional space, then `>` or `/>`.
    if (input.codeUnitAt(end) == _gt) {
      _position = end + 1;
      return xml_events.XmlStartElementEvent(name, attributes, false);
    }
    if (input.codeUnitAt(end) == _slash &&
        end + 1 < input.length &&
        input.codeUnitAt(end + 1) == _gt) {
      _position = end + 2;
      return xml_events.XmlStartElementEvent(name, attributes, true);
    }
    return null;
  }

  xml_events.XmlEndElementEvent? _endElement(int start) {
    final nameEnd = _nameEnd(start + 2);
    if (nameEnd == -1) return null;
    final end = _skipSpace(nameEnd);
    if (end >= _input.length || _input.codeUnitAt(end) != _gt) return null;
    _position = end + 1;
    return xml_events.XmlEndElementEvent(_input.substring(start + 2, nameEnd));
  }

  /// `<?xml version="1.0" ...?>`; other processing instructions delegate.
  xml_events.XmlDeclarationEvent? _declaration(int start) {
    final input = _input;
    if (!input.startsWith('<?xml', start)) return null;
    final afterName = start + 5;
    if (afterName >= input.length) return null;
    final c = input.codeUnitAt(afterName);
    if (!_isSpace(c) && c != _question) return null;
    final attributes = <xml_events.XmlEventAttribute>[];
    final end = _attributes(afterName, attributes);
    if (end == -1 ||
        input.codeUnitAt(end) != _question ||
        end + 1 >= input.length ||
        input.codeUnitAt(end + 1) != _gt) {
      return null;
    }
    _position = end + 2;
    return xml_events.XmlDeclarationEvent(attributes);
  }

  /// Reads ` name="value"` pairs from [position] into [attributes] and
  /// returns the position after them and any trailing space, or -1.
  int _attributes(int position, List<xml_events.XmlEventAttribute> attributes) {
    final input = _input;
    var p = position;
    while (true) {
      final afterSpace = _skipSpace(p);
      if (afterSpace >= input.length) return -1;
      final c = input.codeUnitAt(afterSpace);
      if (c == _gt || c == _slash || c == _question) return afterSpace;
      if (afterSpace == p) return -1; // attributes need a space before them
      final nameEnd = _nameEnd(afterSpace);
      if (nameEnd == -1) return -1;
      var q = _skipSpace(nameEnd);
      if (q >= input.length || input.codeUnitAt(q) != _equals) return -1;
      q = _skipSpace(q + 1);
      if (q >= input.length) return -1;
      final quote = input.codeUnitAt(q);
      if (quote != _quote && quote != _apostrophe) return -1;
      final valueEnd = input.indexOf(quote == _quote ? '"' : "'", q + 1);
      if (valueEnd == -1) return -1;
      attributes.add(xml_events.XmlEventAttribute(
        input.substring(afterSpace, nameEnd),
        _decode(input.substring(q + 1, valueEnd)),
        quote == _quote ? XmlAttributeType.DOUBLE_QUOTE : XmlAttributeType.SINGLE_QUOTE,
      ));
      p = valueEnd + 1;
    }
  }

  /// The end of the XML name starting at [position], or -1 when there is no
  /// valid name there.
  int _nameEnd(int position) {
    final input = _input;
    if (position >= input.length || !_isNameStart(input.codeUnitAt(position))) {
      return -1;
    }
    var p = position + 1;
    while (p < input.length && _isNameChar(input.codeUnitAt(p))) {
      p++;
    }
    return p;
  }

  int _skipSpace(int position) {
    var p = position;
    while (p < _input.length && _isSpace(_input.codeUnitAt(p))) {
      p++;
    }
    return p;
  }

  static String _decode(String raw) =>
      raw.contains('&') ? _mapping.decode(raw) : raw;

  static bool _isSpace(int c) => c == 0x20 || c == 0x09 || c == 0x0A || c == 0x0D;

  // XML NameStartChar / NameChar for ASCII; any non-ASCII character is
  // accepted (spreadsheet XML uses ASCII names).
  static bool _isNameStart(int c) =>
      (c >= 0x61 && c <= 0x7A) || // a-z
      (c >= 0x41 && c <= 0x5A) || // A-Z
      c == 0x5F || // _
      c == 0x3A || // :
      c >= 0xC0;

  static bool _isNameChar(int c) =>
      _isNameStart(c) ||
      (c >= 0x30 && c <= 0x39) || // 0-9
      c == 0x2D || // -
      c == 0x2E || // .
      c == 0xB7;
}
