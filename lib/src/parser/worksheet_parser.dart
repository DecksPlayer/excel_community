part of '../../../excel_community.dart';

/// Parses individual worksheets: sheet data, rows, cells, header/footer,
/// column widths and row heights.
///
/// Separated from [Parser] to keep each file focused on a single concern.
class _WorksheetParser {
  final Excel _excel;
  final Map<String, String> _worksheetTargets;

  _WorksheetParser(this._excel, this._worksheetTargets);

  // ---------------------------------------------------------------------------
  // Public API (called by Parser)
  // ---------------------------------------------------------------------------

  /// Normalizes the sheet after parsing — clears data if the sheet is empty
  /// and refreshes the row/column counts.
  void normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowsAndColumns();
  }

  /// Parses a single `<sheet>` node from the workbook, loads its XML file
  /// and populates the [Sheet] object with all cell data and metadata.
  void parseTable(XmlElement node) {
    final name = node.getAttribute('name')!;
    final target = _worksheetTargets[node.getAttribute('r:id')];

    _excel._sheetMap['$name'] ??= Sheet._(_excel, '$name');
    final sheetObject = _excel._sheetMap['$name']!;

    final file = _excel._archive.findFile('xl/$target')!;
    file.decompress();

    final content = XmlDocument.parse(utf8.decode(file.content));
    final worksheet = content.findElements('worksheet').first;

    // Right-to-left view
    final sheetViews = worksheet.findAllElements('sheetView').toList();
    if (sheetViews.isNotEmpty) {
      final rtl = sheetViews.first.getAttribute('rightToLeft');
      sheetObject.isRTL = rtl != null && rtl == '1';
    }

    final sheetData = worksheet.findElements('sheetData').first;
    _findRows(sheetData).forEach((row) => _parseRow(row, sheetObject, name));

    _parseHeaderFooter(worksheet, sheetObject);
    _parseColWidthsRowHeights(worksheet, sheetObject);

    _excel._sheets[name] = sheetData;
    _excel._xmlFiles['xl/$target'] = content;
    _excel._xmlSheetId[name] = 'xl/$target';

    normalizeTable(sheetObject);
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  void _parseRow(XmlElement node, Sheet sheetObject, String name) {
    final rowIndex = (_getRowNumber(node) ?? -1) - 1;
    if (rowIndex < 0) return;
    _findCells(node).forEach((cell) => _parseCell(cell, sheetObject, rowIndex, name));
  }

  void _parseCell(XmlElement node, Sheet sheetObject, int rowIndex, String name) {
    final columnIndex = _getCellNumber(node);
    if (columnIndex == null) return;

    final s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1);
      } catch (_) {}

      final rC = node.getAttribute('r').toString();
      if (_excel._cellStyleReferenced[name] == null) {
        _excel._cellStyleReferenced[name] = {rC: s};
      } else {
        _excel._cellStyleReferenced[name]![rC] = s;
      }
    }

    CellValue? value;
    final type = node.getAttribute('t');

    switch (type) {
      // Shared string
      case 's':
        final idx = int.parse(_parseXmlValue(node.findElements('v').first).trim());
        value = TextCellValue.span(_excel._sharedStrings.value(idx)!.textSpan);
        break;
      // Boolean
      case 'b':
        value = BoolCellValue(_parseXmlValue(node.findElements('v').first) == '1');
        break;
      // Error / formula string
      case 'e':
      case 'str':
        value = FormulaCellValue(_parseXmlValue(node.findElements('v').first));
        break;
      // Inline string
      case 'inlineStr':
        value = TextCellValue(_parseXmlValue(node.findAllElements('t').first));
        break;
      // Number (default)
      case 'n':
      default:
        final formulaNodes = node.findElements('f');
        if (formulaNodes.isNotEmpty) {
          value = FormulaCellValue(_parseXmlValue(formulaNodes.first));
        } else {
          final vNode = node.findElements('v').firstOrNull;
          if (vNode == null) {
            value = null;
          } else if (s1 != null) {
            final v = _parseXmlValue(vNode);
            final numFmtId = _excel._numFmtIds[s];
            final numFormat =
                _excel._numFormats.getByNumFmtId(numFmtId) ?? NumFormat.standard_0;
            value = numFormat.read(v);
          } else {
            value = NumFormat.defaultNumeric.read(_parseXmlValue(vNode));
          }
        }
    }

    sheetObject.updateCell(
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: _excel._cellStyleList[s],
    );
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements('headerFooter');
    if (results.isEmpty) return;
    sheetObject.headerFooter =
        HeaderFooter.fromXmlElement(results.first);
  }

  void _parseColWidthsRowHeights(XmlElement worksheet, Sheet sheetObject) {
    // --- Default column width / default row height ---
    // Example: <sheetFormatPr defaultColWidth="26.33" defaultRowHeight="13"/>
    final formatPr = worksheet.findAllElements('sheetFormatPr');
    if (formatPr.isNotEmpty) {
      for (final element in formatPr) {
        final colW = double.tryParse(element.getAttribute('defaultColWidth') ?? '');
        final rowH = double.tryParse(element.getAttribute('defaultRowHeight') ?? '');
        if (colW != null && rowH != null) {
          sheetObject._defaultColumnWidth = colW;
          sheetObject._defaultRowHeight = rowH;
        }
      }
    }

    // --- Custom column widths ---
    // Example: <col min="2" max="2" width="71.83" customWidth="1"/>
    final cols = worksheet.findAllElements('col');
    for (final element in cols) {
      final col = int.tryParse(element.getAttribute('min') ?? '');
      final width = double.tryParse(element.getAttribute('width') ?? '');
      if (col != null && width != null) {
        final zeroBasedCol = col - 1;
        if (zeroBasedCol >= 0) {
          sheetObject._columnWidths[zeroBasedCol] = width;
        }
      }
    }

    // --- Custom row heights ---
    // Example: <row r="1" ht="44" customHeight="1"/>
    final rows = worksheet.findAllElements('row');
    for (final element in rows) {
      final row = int.tryParse(element.getAttribute('r') ?? '');
      final height = double.tryParse(element.getAttribute('ht') ?? '');
      if (row != null && height != null) {
        final zeroBasedRow = row - 1;
        if (zeroBasedRow >= 0) {
          sheetObject._rowHeights[zeroBasedRow] = height;
        }
      }
    }
  }
}
