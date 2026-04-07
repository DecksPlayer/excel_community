part of '../../../excel_community.dart';

class _StyleXmlBuilders {
  final Excel _excel;
  final Save _save;

  _StyleXmlBuilders(this._excel, this._save);

  void buildFonts(XmlElement fonts, List<_FontStyle> innerFontStyle) {
    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + innerFontStyle.length}';
    } else {
      fonts.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}'));
    }

    for (var fontStyleElement in innerFontStyle) {
      fonts.children.add(XmlElement(XmlName('font'), [], [
        if (fontStyleElement._fontColorHex != null &&
            fontStyleElement._fontColorHex!.colorHex != "FF000000")
          XmlElement(XmlName('color'), [
            XmlAttribute(
                XmlName('rgb'), fontStyleElement._fontColorHex!.colorHex)
          ], []),

        if (fontStyleElement.isBold) XmlElement(XmlName('b'), [], []),
        if (fontStyleElement.isItalic) XmlElement(XmlName('i'), [], []),
        if (fontStyleElement.isStrikethrough) XmlElement(XmlName('strike'), [], []),

        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline == Underline.Single)
          XmlElement(XmlName('u'), [], []),

        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline != Underline.Single &&
            fontStyleElement.underline == Underline.Double)
          XmlElement(
              XmlName('u'), [XmlAttribute(XmlName('val'), 'double')], []),

        if (fontStyleElement.fontFamily != null &&
            fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
            fontStyleElement.fontFamily != '' &&
            fontStyleElement.fontFamily!.isNotEmpty)
          XmlElement(XmlName('name'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontFamily.toString())
          ], []),

        if (fontStyleElement.fontScheme != FontScheme.Unset)
          XmlElement(XmlName('scheme'), [
            XmlAttribute(
                XmlName('val'),
                switch (fontStyleElement.fontScheme) {
                  FontScheme.Major => "major",
                  _ => "minor"
                })
          ], []),

        if (fontStyleElement.fontSize != null &&
            fontStyleElement.fontSize.toString().isNotEmpty)
          XmlElement(XmlName('sz'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontSize.toString())
          ], []),
      ]));
    }
  }

  void buildFills(XmlElement fills, List<String> innerPatternFill) {
    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value =
          '${_excel._patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}'));
    }

    for (var color in innerPatternFill) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'), [
              XmlAttribute(XmlName('patternType'), 'solid')
            ], [
              XmlElement(XmlName('fgColor'),
                  [XmlAttribute(XmlName('rgb'), color)], []),
              XmlElement(
                  XmlName('bgColor'), [XmlAttribute(XmlName('rgb'), color)], [])
            ])
          ]));
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'),
                [XmlAttribute(XmlName('patternType'), color)], [])
          ]));
        }
      } else {
        _damagedExcel(
            text:
                "Corrupted Styles Found. Can't process further, Open up issue in github.");
      }
    }
  }

  void buildBorders(XmlElement borders, List<_BorderSet> innerBorderSet) {
    var borderAttribute = borders.getAttributeNode('count');

    if (borderAttribute != null) {
      borderAttribute.value =
          '${_excel._borderSetList.length + innerBorderSet.length}';
    } else {
      borders.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._borderSetList.length + innerBorderSet.length}'));
    }

    for (var border in innerBorderSet) {
      var borderElement = XmlElement(XmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes
            .add(XmlAttribute(XmlName('diagonalDown'), '1'));
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes.add(XmlAttribute(XmlName('diagonalUp'), '1'));
      }
      final Map<String, Border> borderMap = {
        'left': border.leftBorder,
        'right': border.rightBorder,
        'top': border.topBorder,
        'bottom': border.bottomBorder,
        'diagonal': border.diagonalBorder,
      };
      for (var key in borderMap.keys) {
        final borderValue = borderMap[key]!;

        final element = XmlElement(XmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes.add(XmlAttribute(XmlName('style'), style.style));
        }
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(XmlElement(
              XmlName('color'), [XmlAttribute(XmlName('rgb'), color)]));
        }
        borderElement.children.add(element);
      }

      borders.children.add(borderElement);
    }
  }

  void buildCellXfs(
    XmlElement celx,
    List<CellStyle> innerCellStyle,
    List<_FontStyle> innerFontStyle,
    List<String> innerPatternFill,
    List<_BorderSet> innerBorderSet,
  ) {
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + innerCellStyle.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._cellStyleList.length + innerCellStyle.length}'));
    }

    for (var cellStyle in innerCellStyle) {
      String backgroundColor = cellStyle.backgroundColor.colorHex;

      _FontStyle fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          strikethrough: cellStyle.isStrikethrough,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily,
          fontScheme: cellStyle.fontScheme);

      int fontId = _fontStyleIndex(_excel._fontStyleList, fs);
      if (fontId == -1) {
        fontId = _fontStyleIndex(innerFontStyle, fs);
        if (fontId != -1) {
          fontId += _excel._fontStyleList.length;
        } else {
          fontId = 0;
        }
      }

      int fillId = _excel._patternFill.indexOf(backgroundColor);
      if (fillId == -1) {
        fillId = innerPatternFill.indexOf(backgroundColor);
        if (fillId != -1) {
          fillId += _excel._patternFill.length;
        } else {
          fillId = 0;
        }
      }

      final bs = _save._createBorderSetFromCellStyle(cellStyle);
      int borderId = _excel._borderSetList.indexOf(bs);
      if (borderId == -1) {
        borderId = innerBorderSet.indexOf(bs);
        if (borderId != -1) {
          borderId += _excel._borderSetList.length;
        } else {
          borderId = 0;
        }
      }

      final numberFormat = cellStyle.numberFormat;
      final int numFmtId = switch (numberFormat) {
        StandardNumFormat() => numberFormat.numFmtId,
        CustomNumFormat() => _excel._numFormats.findOrAdd(numberFormat),
      };

      celx.children.add(XmlElement(XmlName('xf'), [
        XmlAttribute(XmlName('applyFont'), '1'),
        XmlAttribute(XmlName('applyFill'), '1'),
        XmlAttribute(XmlName('applyBorder'), '1'),
        XmlAttribute(XmlName('applyAlignment'), '1'),
        XmlAttribute(XmlName('borderId'), '$borderId'),
        XmlAttribute(XmlName('fillId'), '$fillId'),
        XmlAttribute(XmlName('fontId'), '$fontId'),
        XmlAttribute(XmlName('numFmtId'), numFmtId.toString()),
      ], [
        XmlElement(XmlName('alignment'), [
          XmlAttribute(XmlName('horizontal'),
              cellStyle.horizontalAlignment.toString().split('.').last.toLowerCase()),
          XmlAttribute(XmlName('vertical'),
              cellStyle.verticalAlignment.toString().split('.').last.toLowerCase()),
          XmlAttribute(XmlName('textRotation'), cellStyle.rotation.toString()),
          XmlAttribute(XmlName('wrapText'),
              cellStyle.wrap == TextWrapping.WrapText ? '1' : '0'),
          XmlAttribute(XmlName('shrinkToFit'),
              cellStyle.wrap == TextWrapping.Clip ? '1' : '0'),
        ], []),
      ]));
    }
  }

  void buildNumFmts(XmlDocument styleSheet) {
    final customNumberFormats = _excel._numFormats._map.entries
        .map<MapEntry<int, CustomNumFormat>?>((e) {
          final format = e.value;
          if (format is! CustomNumFormat) {
            return null;
          }
          return MapEntry<int, CustomNumFormat>(e.key, format);
        })
        .nonNulls
        .sorted((a, b) => a.key.compareTo(b.key));

    if (customNumberFormats.isNotEmpty) {
      var numFmtsElement = styleSheet
          .findAllElements('numFmts')
          .whereType<XmlElement>()
          .firstOrNull;
      int count;
      if (numFmtsElement == null) {
        numFmtsElement = XmlElement(XmlName('numFmts'));

        styleSheet
            .findElements('styleSheet')
            .first
            .children
            .insert(0, numFmtsElement);
      }
      count = int.parse(numFmtsElement.getAttribute('count') ?? '0');

      for (var numFormat in customNumberFormats) {
        final numFmtIdString = numFormat.key.toString();
        final formatCode = numFormat.value.formatCode;
        var numFmtElement = numFmtsElement.children
            .whereType<XmlElement>()
            .firstWhereOrNull((node) =>
                node.name.local == 'numFmt' &&
                node.getAttribute('numFmtId') == numFmtIdString);
        if (numFmtElement == null) {
          numFmtElement = XmlElement(
              XmlName('numFmt'),
              [
                XmlAttribute(XmlName('numFmtId'), numFmtIdString),
                XmlAttribute(XmlName('formatCode'), formatCode),
              ],
              [],
              true);
          numFmtsElement.children.add(numFmtElement);
          count++;
        } else if ((numFmtElement.getAttribute('formatCode') ?? '') !=
            formatCode) {
          numFmtElement.setAttribute('formatCode', formatCode);
        }
      }

      numFmtsElement.setAttribute('count', count.toString());
    }
  }
}
