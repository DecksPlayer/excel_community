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
      fonts.attributes.add(XmlAttribute(_xmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}'));
    }

    for (var fontStyleElement in innerFontStyle) {
      fonts.children.add(XmlElement(_xmlName('font'), [], [
        if (fontStyleElement.fontFamily != null &&
            fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
            fontStyleElement.fontFamily != '' &&
            fontStyleElement.fontFamily!.isNotEmpty)
          XmlElement(_xmlName('name'), [
            XmlAttribute(
                _xmlName('val'), fontStyleElement.fontFamily.toString())
          ], []),
        if (fontStyleElement.isBold) XmlElement(_xmlName('b'), [], []),
        if (fontStyleElement.isItalic) XmlElement(_xmlName('i'), [], []),
        if (fontStyleElement.isStrikethrough)
          XmlElement(_xmlName('strike'), [], []),
        if (fontStyleElement._fontColorHex != null &&
            fontStyleElement._fontColorHex!.colorHex != "FF000000")
          XmlElement(_xmlName('color'), [
            XmlAttribute(
                _xmlName('rgb'), fontStyleElement._fontColorHex!.colorHex)
          ], []),
        if (fontStyleElement.fontSize != null &&
            fontStyleElement.fontSize.toString().isNotEmpty)
          XmlElement(_xmlName('sz'), [
            XmlAttribute(
                _xmlName('val'), fontStyleElement.fontSize.toString())
          ], []),
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline == Underline.Single)
          XmlElement(_xmlName('u'), [], []),
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline != Underline.Single &&
            fontStyleElement.underline == Underline.Double)
          XmlElement(_xmlName('u'),
              [XmlAttribute(_xmlName('val'), 'double')], []),
        if (fontStyleElement.fontScheme != FontScheme.Unset)
          XmlElement(_xmlName('scheme'), [
            XmlAttribute(
                _xmlName('val'),
                switch (fontStyleElement.fontScheme) {
                  FontScheme.Major => "major",
                  _ => "minor"
                })
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
      fills.attributes.add(XmlAttribute(_xmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}'));
    }

    // One <fill> per entry, always: fillId is the entry's position, so a
    // skipped entry would shift every fill after it.
    for (var color in innerPatternFill) {
      if (color.length >= 2) {
        if (color == "none" || color == "gray125" || color == "lightGray") {
          fills.children.add(XmlElement(_xmlName('fill'), [], [
            XmlElement(_xmlName('patternFill'),
                [XmlAttribute(_xmlName('patternType'), color)], [])
          ]));
        } else {
          // An ARGB color, opaque (FF...) or not (e.g. ExcelColor.black12).
          fills.children.add(XmlElement(_xmlName('fill'), [], [
            XmlElement(_xmlName('patternFill'), [
              XmlAttribute(_xmlName('patternType'), 'solid')
            ], [
              XmlElement(_xmlName('fgColor'),
                  [XmlAttribute(_xmlName('rgb'), color)], []),
              XmlElement(_xmlName('bgColor'),
                  [XmlAttribute(_xmlName('rgb'), color)], [])
            ])
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
      borders.attributes.add(XmlAttribute(_xmlName('count'),
          '${_excel._borderSetList.length + innerBorderSet.length}'));
    }

    for (var border in innerBorderSet) {
      var borderElement = XmlElement(_xmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes
            .add(XmlAttribute(_xmlName('diagonalDown'), '1'));
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes
            .add(XmlAttribute(_xmlName('diagonalUp'), '1'));
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

        final element = XmlElement(_xmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes
              .add(XmlAttribute(_xmlName('style'), style.style));
        }
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(XmlElement(_xmlName('color'),
              [XmlAttribute(_xmlName('rgb'), color)]));
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
      celx.attributes.add(XmlAttribute(_xmlName('count'),
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

      final List<XmlAttribute> protectionAttrs = [];
      if (cellStyle.locked != null) {
        protectionAttrs.add(XmlAttribute(_xmlName('locked'), cellStyle.locked! ? 'true' : 'false'));
      }
      if (cellStyle.hidden != null) {
        protectionAttrs.add(XmlAttribute(_xmlName('hidden'), cellStyle.hidden! ? 'true' : 'false'));
      }

      final List<XmlAttribute> xfAttributes = [
        XmlAttribute(_xmlName('applyFont'), '1'),
        XmlAttribute(_xmlName('applyFill'), '1'),
        XmlAttribute(_xmlName('applyBorder'), '1'),
        XmlAttribute(_xmlName('applyAlignment'), '1'),
        if (protectionAttrs.isNotEmpty)
          XmlAttribute(_xmlName('applyProtection'), '1'),
        XmlAttribute(_xmlName('borderId'), '$borderId'),
        XmlAttribute(_xmlName('fillId'), '$fillId'),
        XmlAttribute(_xmlName('fontId'), '$fontId'),
        XmlAttribute(_xmlName('numFmtId'), numFmtId.toString()),
        XmlAttribute(_xmlName('xfId'), '0'),
      ];

      celx.children.add(XmlElement(_xmlName('xf'), xfAttributes, [
        XmlElement(_xmlName('alignment'), [
          XmlAttribute(
              _xmlName('horizontal'),
              cellStyle.horizontalAlignment
                  .toString()
                  .split('.')
                  .last
                  .toLowerCase()),
          XmlAttribute(
              _xmlName('vertical'),
              cellStyle.verticalAlignment
                  .toString()
                  .split('.')
                  .last
                  .toLowerCase()),
          XmlAttribute(
              _xmlName('textRotation'), cellStyle.rotation.toString()),
          XmlAttribute(_xmlName('wrapText'),
              cellStyle.wrap == TextWrapping.WrapText ? '1' : '0'),
          XmlAttribute(_xmlName('shrinkToFit'),
              cellStyle.wrap == TextWrapping.Clip ? '1' : '0'),
        ], []),
        if (protectionAttrs.isNotEmpty)
          XmlElement(_xmlName('protection'), protectionAttrs, []),
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
        numFmtsElement = XmlElement(_xmlName('numFmts'));

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
              _xmlName('numFmt'),
              [
                XmlAttribute(_xmlName('numFmtId'), numFmtIdString),
                XmlAttribute(_xmlName('formatCode'), formatCode),
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

  void buildDxfs(
    XmlDocument styleSheetDoc, [
    List<DifferentialStyle>? innerDxfList,
  ]) {
    if (_excel._dxfList.isEmpty) return;

    final styleSheet = styleSheetDoc.findElements('styleSheet').first;
    var dxfsElement = styleSheet
        .findAllElements('dxfs')
        .whereType<XmlElement>()
        .firstOrNull;

    if (dxfsElement == null) {
      dxfsElement = XmlElement(_xmlName('dxfs'));
      int insertIndex = styleSheet.children.length;
      final tableStylesIdx = styleSheet.children.indexWhere((node) =>
          node is XmlElement &&
          (node.name.local == 'tableStyles' || node.name.local == 'extLst'));
      if (tableStylesIdx != -1) {
        insertIndex = tableStylesIdx;
      }
      styleSheet.children.insert(insertIndex, dxfsElement);
    }

    dxfsElement.children.clear();
    dxfsElement.setAttribute('count', '${_excel._dxfList.length}');

    for (final dxf in _excel._dxfList) {
      final dxfElement = XmlElement(_xmlName('dxf'));

      final fontChildren = <XmlElement>[];
      if (dxf.bold != null) {
        fontChildren.add(XmlElement(_xmlName('b'), [
          if (!dxf.bold!) XmlAttribute(_xmlName('val'), 'false'),
        ]));
      }
      if (dxf.italic != null) {
        fontChildren.add(XmlElement(_xmlName('i'), [
          if (!dxf.italic!) XmlAttribute(_xmlName('val'), 'false'),
        ]));
      }
      if (dxf.strikethrough != null) {
        fontChildren.add(XmlElement(_xmlName('strike'), [
          if (!dxf.strikethrough!) XmlAttribute(_xmlName('val'), 'false'),
        ]));
      }
      if (dxf.underline != null && dxf.underline != Underline.None) {
        final val = dxf.underline == Underline.Double ? 'double' : 'single';
        fontChildren.add(XmlElement(_xmlName('u'), [
          XmlAttribute(_xmlName('val'), val),
        ]));
      }
      if (dxf.fontColor != null) {
        fontChildren.add(XmlElement(_xmlName('color'), [
          XmlAttribute(_xmlName('rgb'), _cleanHex(dxf.fontColor!.colorHex)),
        ]));
      }

      if (fontChildren.isNotEmpty) {
        dxfElement.children.add(XmlElement(_xmlName('font'), [], fontChildren));
      }

      if (dxf.backgroundColor != null) {
        final hex = _cleanHex(dxf.backgroundColor!.colorHex);
        final patternFill = XmlElement(_xmlName('patternFill'), [
          XmlAttribute(_xmlName('patternType'), 'solid'),
        ], [
          XmlElement(_xmlName('fgColor'), [
            XmlAttribute(_xmlName('rgb'), hex),
          ]),
          XmlElement(_xmlName('bgColor'), [
            XmlAttribute(_xmlName('rgb'), hex),
          ]),
        ]);
        dxfElement.children.add(XmlElement(_xmlName('fill'), [], [patternFill]));
      }

      dxfsElement.children.add(dxfElement);
    }
  }

  String _cleanHex(String hex) {
    var clean = hex.replaceAll('#', '').trim().toUpperCase();
    if (clean.length == 6) {
      clean = 'FF$clean';
    }
    return clean;
  }
}
