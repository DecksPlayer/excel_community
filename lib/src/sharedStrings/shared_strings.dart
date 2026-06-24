part of '../../excel_community.dart';

class _SharedStringsMaintainer {
  final Map<String, SharedString> _mapString = <String, SharedString>{};
  final List<SharedString> _list = <SharedString>[];
  int _index = 0;

  _SharedStringsMaintainer._();

  SharedString? tryFind(String xmlKey) {
    return _mapString[xmlKey];
  }

  SharedString addFromString(String val) {
    final newSharedString = SharedString.fromPlainString(val);
    add(newSharedString, newSharedString._xmlString);
    return newSharedString;
  }

  void add(SharedString val, String key) {
    if (val.index != -1) {
      val.count++;
    } else {
      val.index = _index++;
      val.count = 1;
      _mapString[key] = val;
      _list.add(val);
    }
  }

  int indexOf(SharedString val) {
    return val.index;
  }

  SharedString? value(int i) {
    if (i < _list.length) {
      return _list[i];
    } else {
      return null;
    }
  }

  void clear() {
    _index = 0;
    _list.clear();
    _mapString.clear();
  }
}

class SharedString {
  final XmlElement? _node;
  final String _xmlString;
  final String _stringValue;
  final int _hashCode;
  TextSpan? _textSpan;
  int index = -1;
  int count = 1;

  SharedString({required XmlElement node, String? stringValue})
      : this._internal(
          node: node,
          xmlString: node.toXmlString(),
          stringValue: stringValue ?? _getRawStringValue(node),
        );

  SharedString.fromPlainString(String val)
      : _node = null,
        _xmlString = '<si><t xml:space="preserve">${_escapeXml(val)}</t></si>',
        _stringValue = val,
        _textSpan = null,
        _hashCode = val.hashCode;

  SharedString.fromTextSpan(TextSpan span)
      : _node = null,
        _stringValue = span.toString(),
        _textSpan = span,
        _hashCode = span.hashCode,
        _xmlString = _textSpanToXml(span);

  SharedString._internal({
    required XmlElement node,
    required String xmlString,
    required String stringValue,
  })  : _node = node,
        _xmlString = xmlString,
        _stringValue = stringValue,
        _hashCode = xmlString.hashCode,
        _textSpan = null;

  static String _getRawStringValue(XmlElement node) {
    var buffer = StringBuffer();
    node.findAllElements('t').forEach((child) {
      if (child.parentElement == null ||
          child.parentElement!.name.local != 'rPh') {
        buffer.write(_parseXmlValue(child));
      }
    });
    return buffer.toString();
  }

  @override
  String toString() {
    assert(false,
        'prefer stringValue over SharedString.toString() in development');
    return stringValue;
  }

  TextSpan get textSpan {
    if (_textSpan != null) {
      return _textSpan!;
    }
    if (_node == null) {
      _textSpan = TextSpan(text: _stringValue);
      return _textSpan!;
    }
    // Parse an OOXML boolean attribute per ECMA-376 §18.8.2 — the `val` on
    // <b/>, <i/>, etc. is a W3C XSD boolean (lexical space: true|false|1|0).
    // When `val` is omitted the element's presence means on (spec default).
    // The previous implementation used `bool.tryParse`, which only accepts
    // the literal strings "true"/"false", so <b val="0"/> was misread as on.
    bool getBool(XmlElement element) {
      final val = element.getAttribute('val');
      if (val == null) return true;
      final v = val.toLowerCase();
      if (v == 'false' || v == 'f' || v == '0' || v == 'off') return false;
      return true;
    }

    int getDouble(XmlElement element) {
      // Should be double
      return double.parse(element.getAttribute('val')!).toInt();
    }

    String? text;
    List<TextSpan>? children;

    /// SharedStringItem
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem?view=openxml-3.0.1
    assert(_node.name.local == 'si'); //18.4.8 si (String Item)

    for (final child in _node.children.whereType<XmlElement>()) {
      switch (child.name.local) {
        /// Text
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1
        case 't': //18.4.12 t (Text)
          text = (text ?? '') + child.innerText;
          break;

        /// Rich Text Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1
        case 'r': //18.4.4 r (Rich Text Run)
          var style = CellStyle();
          for (final runChild in child.children.whereType<XmlElement>()) {
            switch (runChild.name.local) {
              /// RunProperties
              /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runproperties?view=openxml-3.0.1
              case 'rPr':
                for (final runProperty
                    in runChild.children.whereType<XmlElement>()) {
                  switch (runProperty.name.local) {
                    case 'b': //18.8.2 b (Bold)
                      style = style.copyWith(boldVal: getBool(runProperty));
                      break;
                    case 'i': //18.8.26 i (Italic)
                      style = style.copyWith(italicVal: getBool(runProperty));
                      break;
                    case 'u': //18.4.13 u (Underline)
                      // Per ECMA-376 §18.4.13, <u val="none"/> explicitly
                      // disables underline (ST_UnderlineValues includes
                      // `none`). The previous implementation treated any
                      // non-"double" value as single underline.
                      final uVal = runProperty.getAttribute('val');
                      if (uVal == 'none') break;
                      style = style.copyWith(
                          underlineVal: uVal == 'double'
                              ? Underline.Double
                              : Underline.Single);
                      break;
                    case 'sz': //18.4.11 sz (Font Size)
                      style =
                          style.copyWith(fontSizeVal: getDouble(runProperty));
                      break;
                    case 'rFont': //18.4.5 rFont (Font)
                      style = style.copyWith(
                          fontFamilyVal: runProperty.getAttribute('val'));
                      break;
                    case 'color': //18.3.1.15 color (Data Bar Color)
                      style = style.copyWith(
                          fontColorHexVal:
                              runProperty.getAttribute('rgb')?.excelColor);
                      break;
                  }
                }
                break;

              /// Text
              case 't': //18.4.12 t (Text)
                if (children == null) children = [];
                children.add(TextSpan(text: runChild.innerText, style: style));
                break;
            }
          }
          break;

        /// Phonetic Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticrun?view=openxml-3.0.1
        case 'rPh': //18.4.6 rPh (Phonetic Run)
          break;
      }
    }

    return TextSpan(text: text, children: children);
  }

  String get stringValue => _stringValue;

  @override
  int get hashCode => _hashCode;

  @override
  bool operator ==(Object other) {
    return other is SharedString &&
        other.hashCode == _hashCode &&
        other._xmlString == _xmlString;
  }

  bool matches(String value) {
    return value.isNotEmpty && value == _stringValue;
  }
}

class TextSpan {
  final String? text;
  final List<TextSpan>? children;
  final CellStyle? style;

  const TextSpan({this.children, this.text, this.style});

  @override
  String toString() {
    String r = '';
    if (text != null) r += text!;
    if (children != null) r += children!.join();
    return r;
  }

  @override
  operator ==(Object other) {
    if (identical(this, other)) return true;
    if (other.runtimeType != runtimeType) return false;
    return other is TextSpan &&
        other.text == text &&
        other.style == style &&
        ListEquality().equals(other.children, children);
  }

  @override
  int get hashCode =>
      Object.hash(text, style, Object.hashAll(children ?? const []));
}

String _textSpanToXml(TextSpan span) {
  final hasStyle = span.style != null && _hasAnyStyle(span.style!);
  final hasChildren = span.children != null && span.children!.isNotEmpty;
  if (!hasStyle && !hasChildren) {
    return '<si><t xml:space="preserve">${_escapeXml(span.text ?? "")}</t></si>';
  }
  final sb = StringBuffer();
  sb.write('<si>');
  _spanToXmlHelper(span, sb);
  sb.write('</si>');
  return sb.toString();
}

void _spanToXmlHelper(TextSpan span, StringBuffer sb) {
  final textVal = span.text ?? '';
  final hasStyle = span.style != null && _hasAnyStyle(span.style!);

  if (textVal.isNotEmpty) {
    if (hasStyle) {
      sb.write('<r>');
      sb.write('<rPr>');
      sb.write(_styleToRPr(span.style));
      sb.write('</rPr>');
      sb.write('<t xml:space="preserve">${_escapeXml(textVal)}</t>');
      sb.write('</r>');
    } else {
      sb.write('<r><t xml:space="preserve">${_escapeXml(textVal)}</t></r>');
    }
  }

  if (span.children != null) {
    for (final child in span.children!) {
      _spanToXmlHelper(child, sb);
    }
  }
}

String _styleToRPr(CellStyle? style) {
  if (style == null) return '';
  final sb = StringBuffer();
  if (style.fontFamily != null &&
      style.fontFamily!.isNotEmpty &&
      style.fontFamily!.toLowerCase() != 'null') {
    sb.write('<rFont val="${_escapeXml(style.fontFamily!)}"/>');
  }
  if (style.isBold) {
    sb.write('<b/>');
  }
  if (style.isItalic) {
    sb.write('<i/>');
  }
  if (style.isStrikethrough) {
    sb.write('<strike/>');
  }
  if (style.fontColor != ExcelColor.none &&
      style.fontColor.colorHex != 'FF000000') {
    sb.write('<color rgb="${style.fontColor.colorHex}"/>');
  }
  if (style.fontSize != null && style.fontSize! > 0) {
    sb.write('<sz val="${style.fontSize}"/>');
  }
  if (style.underline != Underline.None) {
    if (style.underline == Underline.Single) {
      sb.write('<u/>');
    } else if (style.underline == Underline.Double) {
      sb.write('<u val="double"/>');
    }
  }
  return sb.toString();
}

bool _hasAnyStyle(CellStyle style) {
  return style.isBold ||
      style.isItalic ||
      style.isStrikethrough ||
      style.underline != Underline.None ||
      (style.fontSize != null && style.fontSize! > 0) ||
      (style.fontFamily != null &&
          style.fontFamily!.isNotEmpty &&
          style.fontFamily!.toLowerCase() != 'null') ||
      (style.fontColor != ExcelColor.none &&
          style.fontColor.colorHex != 'FF000000');
}
