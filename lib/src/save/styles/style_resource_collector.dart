part of '../../../excel_community.dart';

class _StyleResources {
  final List<CellStyle> innerCellStyle = [];
  final List<String> innerPatternFill = [];
  final List<_FontStyle> innerFontStyle = [];
  final List<_BorderSet> innerBorderSet = [];
  final List<DifferentialStyle> innerDxfList = [];
}

class _StyleResourceCollector {
  final Excel _excel;
  final Save _save;

  _StyleResourceCollector(this._excel, this._save);

  _StyleResources collect() {
    final resources = _StyleResources();

    final uniqueStyles = <CellStyle>{};
    final lastStyles = List<CellStyle?>.filled(8, null);
    int lastIdx = 0;

    // 1. Gather all unique CellStyle objects from all sheets
    _excel._sheetMap.forEach((sheetName, sheet) {
      sheet._sheetData.forEach((_, columnMap) {
        columnMap.forEach((_, dataObject) {
          final style = dataObject._cellStyle;
          if (style != null) {
            for (int i = 0; i < 8; i++) {
              if (identical(lastStyles[i], style)) {
                return;
              }
            }
            lastStyles[lastIdx] = style;
            lastIdx = (lastIdx + 1) & 7;

            if (uniqueStyles.add(style)) {
              resources.innerCellStyle.add(style);
            }
          }
        });
      });

      // Gather differential styles from conditional formatting rules
      for (final group in sheet._conditionalFormattings) {
        for (final rule in group.rules) {
          final dxf = rule.style;
          if (!dxf.isEmpty) {
            if (!_excel._dxfList.contains(dxf) &&
                !resources.innerDxfList.contains(dxf)) {
              resources.innerDxfList.add(dxf);
            }
          }
        }
      }
    });

    // 2. Extract unique fonts, fills, and borders from collected CellStyles
    for (var cellStyle in resources.innerCellStyle) {
      // Font extraction
      final fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          strikethrough: cellStyle.isStrikethrough,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily,
          fontScheme: cellStyle.fontScheme);

      if (_fontStyleIndex(_excel._fontStyleList, fs) == -1 &&
          _fontStyleIndex(resources.innerFontStyle, fs) == -1) {
        resources.innerFontStyle.add(fs);
      }

      // Fill extraction
      final backgroundColor = cellStyle.backgroundColor.colorHex;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !resources.innerPatternFill.contains(backgroundColor)) {
        resources.innerPatternFill.add(backgroundColor);
      }

      // Border extraction
      final bs = _save._createBorderSetFromCellStyle(cellStyle);
      if (!_excel._borderSetList.contains(bs) &&
          !resources.innerBorderSet.contains(bs)) {
        resources.innerBorderSet.add(bs);
      }
    }

    return resources;
  }
}
