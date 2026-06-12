part of '../../../excel_community.dart';

/// Parses styles.xml: fonts, fills, borders, number formats and cell styles.
///
/// Separated from [Parser] to keep each file focused on a single concern.
class _StylesParser {
  final Excel _excel;

  _StylesParser(this._excel);

  /// Entry point — parses the styles target file and populates all style lists
  /// on the [Excel] instance.
  void parse(String stylesTarget) {
    var styles = _excel._archive.findFile('xl/$stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._cellStyleList = <CellStyle>[];
      _excel._borderSetList = <_BorderSet>[];

      Iterable<XmlElement> fontList = document.findAllElements('font');
      for (final font in fontList) {
        _excel._fontStyleList.add(_parseFontStyle(font));
      }

      _parseFills(document);
      _parseBorders(document);
      _parseNumFmts(document);
      _parseCellXfs(document, fontList);
    } else {
      _damagedExcel(text: 'styles');
    }
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  void _parseFills(XmlDocument document) {
    document.findAllElements('patternFill').forEach((node) {
      String patternType = node.getAttribute('patternType') ?? '';
      if (node.children.isNotEmpty) {
        node.findElements('fgColor').forEach((child) {
          final rgb = child.getAttribute('rgb') ?? '';
          _excel._patternFill.add(rgb);
        });
      } else {
        _excel._patternFill.add(patternType);
      }
    });
  }

  void _parseBorders(XmlDocument document) {
    document.findAllElements('border').forEach((node) {
      final diagonalUp = !['0', 'false', null]
          .contains(node.getAttribute('diagonalUp')?.trim());
      final diagonalDown = !['0', 'false', null]
          .contains(node.getAttribute('diagonalDown')?.trim());

      const List<String> borderSides = [
        'left',
        'right',
        'top',
        'bottom',
        'diagonal',
      ];

      Map<String, Border> borderElements = {};
      for (var side in borderSides) {
        XmlElement? element;
        try {
          element = node.findElements(side).single;
        } on StateError catch (_) {
          // Either there is no element, or there are too many ones.
          // Silently ignore this element.
        }

        final borderStyleAttribute = element?.getAttribute('style')?.trim();
        final borderStyle = borderStyleAttribute != null
            ? getBorderStyleByName(borderStyleAttribute)
            : null;

        String? borderColorHex;
        try {
          final color = element?.findElements('color').single;
          borderColorHex = color?.getAttribute('rgb')?.trim();
        } on StateError catch (_) {}

        borderElements[side] = Border(
          borderStyle: borderStyle,
          borderColorHex: borderColorHex?.excelColor,
        );
      }

      _excel._borderSetList.add(_BorderSet(
        leftBorder: borderElements['left']!,
        rightBorder: borderElements['right']!,
        topBorder: borderElements['top']!,
        bottomBorder: borderElements['bottom']!,
        diagonalBorder: borderElements['diagonal']!,
        diagonalBorderDown: diagonalDown,
        diagonalBorderUp: diagonalUp,
      ));
    });
  }

  void _parseNumFmts(XmlDocument document) {
    document.findAllElements('numFmts').forEach((node1) {
      node1.findAllElements('numFmt').forEach((node) {
        final numFmtId = int.parse(node.getAttribute('numFmtId')!);
        final formatCode = node.getAttribute('formatCode')!;
        // Register all explicitly declared formats. Excel can include
        // built-in IDs (0–163) in numFmts to override them for this file.
        _excel._numFormats
            .add(numFmtId, NumFormat.custom(formatCode: formatCode));
      });
    });
  }

  void _parseCellXfs(XmlDocument document, Iterable<XmlElement> fontList) {
    document.findAllElements('cellXfs').forEach((node1) {
      node1.findAllElements('xf').forEach((node) {
        final numFmtId = _getFontIndex(node, 'numFmtId');
        _excel._numFmtIds.add(numFmtId);

        String fontColor = ExcelColor.black.colorHex,
            backgroundColor = ExcelColor.none.colorHex;
        String? fontFamily;
        FontScheme fontScheme = FontScheme.Unset;
        _BorderSet? borderSet;
        int fontSize = 12;
        bool isBold = false, isItalic = false, isStrikethrough = false;
        Underline underline = Underline.None;
        HorizontalAlign horizontalAlign = HorizontalAlign.Left;
        VerticalAlign verticalAlign = VerticalAlign.Bottom;
        TextWrapping? textWrapping;
        int rotation = 0;

        final fontId = _getFontIndex(node, 'fontId');
        if (fontId < fontList.length) {
          final fontStyle = _parseFontStyle(fontList.elementAt(fontId));
          fontColor = fontStyle.fontColor.colorHex;
          fontSize = fontStyle.fontSize ?? fontSize;
          isBold = fontStyle.isBold;
          isItalic = fontStyle.isItalic;
          isStrikethrough = fontStyle.isStrikethrough;
          underline = fontStyle.underline;
          fontFamily = fontStyle.fontFamily;
          fontScheme = fontStyle.fontScheme;
        }

        final fillId = _getFontIndex(node, 'fillId');
        if (fillId < _excel._patternFill.length) {
          backgroundColor = _excel._patternFill[fillId];
        }

        final borderId = _getFontIndex(node, 'borderId');
        if (borderId < _excel._borderSetList.length) {
          borderSet = _excel._borderSetList[borderId];
        }

        if (node.children.isNotEmpty) {
          node.findElements('alignment').forEach((child) {
            if (_getFontIndex(child, 'wrapText') == 1) {
              textWrapping = TextWrapping.WrapText;
            } else if (_getFontIndex(child, 'shrinkToFit') == 1) {
              textWrapping = TextWrapping.Clip;
            }

            final vertical =
                child.getAttribute('vertical') ?? node.getAttribute('vertical');
            if (vertical != null) {
              if (vertical == 'top') {
                verticalAlign = VerticalAlign.Top;
              } else if (vertical == 'center') {
                verticalAlign = VerticalAlign.Center;
              }
            }

            final horizontal = child.getAttribute('horizontal') ??
                node.getAttribute('horizontal');
            if (horizontal != null) {
              if (horizontal == 'center') {
                horizontalAlign = HorizontalAlign.Center;
              } else if (horizontal == 'right') {
                horizontalAlign = HorizontalAlign.Right;
              }
            }

            final rotationString = child.getAttribute('textRotation') ??
                node.getAttribute('textRotation');
            if (rotationString != null) {
              rotation = (double.tryParse(rotationString) ?? 0.0).floor();
            }
          });
        }

        // getByNumFmtId returns null only for truly unknown custom IDs.
        // Built-in IDs (0–163) always get a fallback via NumFormatMaintainer.
        final numFormat =
            _excel._numFormats.getByNumFmtId(numFmtId) ?? NumFormat.standard_0;

        _excel._cellStyleList.add(CellStyle(
          fontColorHex: fontColor.excelColor,
          fontFamily: fontFamily,
          fontScheme: fontScheme,
          fontSize: fontSize,
          bold: isBold,
          italic: isItalic,
          strikethrough: isStrikethrough,
          underline: underline,
          backgroundColorHex:
              backgroundColor == 'none' || backgroundColor.isEmpty
                  ? ExcelColor.none
                  : backgroundColor.excelColor,
          horizontalAlign: horizontalAlign,
          verticalAlign: verticalAlign,
          textWrapping: textWrapping,
          rotation: rotation,
          leftBorder: borderSet?.leftBorder,
          rightBorder: borderSet?.rightBorder,
          topBorder: borderSet?.topBorder,
          bottomBorder: borderSet?.bottomBorder,
          diagonalBorder: borderSet?.diagonalBorder,
          diagonalBorderUp: borderSet?.diagonalBorderUp ?? false,
          diagonalBorderDown: borderSet?.diagonalBorderDown ?? false,
          numberFormat: numFormat,
        ));
      });
    });
  }

  /// Returns the value of an XML child element or attribute, or null.
  ///
  /// Returns `true` when the child exists but [attribute] is not specified
  /// (useful for presence-only elements like `<b>`, `<i>`).
  dynamic _nodeChildren(XmlElement node, String child, {dynamic attribute}) {
    final elements = node.findElements(child);
    if (elements.isNotEmpty) {
      if (attribute != null) {
        return elements.first.getAttribute(attribute as String);
      }
      return true;
    }
    return null;
  }

  /// Parses an integer index from an XML attribute, defaulting to 0.
  /// Treats `"true"` as 1 (used for boolean-style attributes in OOXML).
  int _getFontIndex(XmlElement node, String text) {
    final raw = node.getAttribute(text)?.trim();
    if (raw != null) {
      final parsed = int.tryParse(raw);
      if (parsed != null) return parsed;
      if (raw.toLowerCase() == 'true') return 1;
    }
    return 0;
  }

  _FontStyle _parseFontStyle(XmlElement font) {
    final fontStyle = _FontStyle();

    final color = _nodeChildren(font, 'color', attribute: 'rgb');
    if (color != null && color is! bool) {
      fontStyle._fontColorHex = color.toString().excelColor;
    }

    final size = _nodeChildren(font, 'sz', attribute: 'val');
    if (size != null) {
      fontStyle.fontSize = double.parse(size as String).round();
    }

    final bold = _nodeChildren(font, 'b');
    if (bold != null && bold is bool && bold) fontStyle.isBold = true;

    final italic = _nodeChildren(font, 'i');
    if (italic != null && italic is bool && italic) fontStyle.isItalic = true;

    final strike = _nodeChildren(font, 'strike');
    if (strike != null && strike is bool && strike) {
      fontStyle.isStrikethrough = true;
    }

    final underlineValue = _nodeChildren(font, 'u', attribute: 'val');
    if (underlineValue != null && underlineValue != true) {
      if ((underlineValue as String).toLowerCase() == 'double') {
        fontStyle.underline = Underline.Double;
      }
    } else {
      final underlineElement = _nodeChildren(font, 'u');
      if (underlineElement != null && underlineElement == true) {
        fontStyle.underline = Underline.Single;
      }
    }

    final family = _nodeChildren(font, 'name', attribute: 'val');
    if (family != null && family != true) {
      fontStyle.fontFamily = family as String;
    }

    final scheme = _nodeChildren(font, 'scheme', attribute: 'val');
    if (scheme != null) {
      fontStyle.fontScheme =
          scheme == 'major' ? FontScheme.Major : FontScheme.Minor;
    }

    return fontStyle;
  }
}
