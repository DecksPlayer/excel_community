part of '../../../excel_community.dart';

/// Parses individual worksheets using SAX/event-based parsing to achieve
/// high performance and a low memory footprint.
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

  /// Parses a single `<sheet>` node from the workbook, loads its XML file,
  /// parses it using SAX event streaming, and populates the [Sheet] object.
  void parseTable(XmlElement node) {
    final name = node.getAttribute('name')!;
    final target = _worksheetTargets[node.getAttribute('r:id')];
    if (target == null) {
      throw ArgumentError('Worksheet target not found for relationship ID ${node.getAttribute('r:id')}');
    }

    String path = target;
    if (path.startsWith('/')) {
      path = path.substring(1);
    } else if (!path.startsWith('xl/')) {
      path = 'xl/$path';
    }

    _excel._sheetMap['$name'] ??= Sheet._(_excel, '$name');
    final sheetObject = _excel._sheetMap['$name']!;

    final file = _excel._archive.findFile(path)!;
    file.decompress();

    final contentString = utf8.decode(file.content);
    _excel._sheetXmls[path] = contentString;
    _excel._xmlSheetId[name] = path;

    final events = xml_events.parseEvents(contentString);

    List<xml_events.XmlEvent>? currentCellEvents;
    List<xml_events.XmlEvent>? currentHeaderFooterEvents;
    int? currentWorksheetRowIndex;

    for (final event in events) {
      if (event is xml_events.XmlStartElementEvent) {
        final tagName = event.name;

        if (tagName == 'sheetView' || tagName.endsWith(':sheetView')) {
          final rtl = _getAttr(event, 'rightToLeft');
          sheetObject.isRTL = rtl == '1';
        } else if (tagName == 'sheetFormatPr' || tagName.endsWith(':sheetFormatPr')) {
          final colW = double.tryParse(_getAttr(event, 'defaultColWidth') ?? '');
          final rowH = double.tryParse(_getAttr(event, 'defaultRowHeight') ?? '');
          if (colW != null && rowH != null) {
            sheetObject._defaultColumnWidth = colW;
            sheetObject._defaultRowHeight = rowH;
          }
        } else if (tagName == 'col' || tagName.endsWith(':col')) {
          final min = int.tryParse(_getAttr(event, 'min') ?? '');
          final maxVal = int.tryParse(_getAttr(event, 'max') ?? '');
          final width = double.tryParse(_getAttr(event, 'width') ?? '');
          if (min != null && width != null) {
            final end = maxVal ?? min;
            for (int col = min; col <= end; col++) {
              final zeroBasedCol = col - 1;
              if (zeroBasedCol >= 0) {
                sheetObject._columnWidths[zeroBasedCol] = width;
              }
            }
          }
        } else if (tagName == 'row' || tagName.endsWith(':row')) {
          final rowNum = int.tryParse(_getAttr(event, 'r') ?? '');
          if (rowNum != null) {
            currentWorksheetRowIndex = rowNum - 1;
            final height = double.tryParse(_getAttr(event, 'ht') ?? '');
            if (height != null && currentWorksheetRowIndex >= 0) {
              sheetObject._rowHeights[currentWorksheetRowIndex] = height;
            }
          }
        } else if (tagName == 'c' || tagName.endsWith(':c')) {
          currentCellEvents = [event];
        } else if (tagName == 'headerFooter' || tagName.endsWith(':headerFooter')) {
          currentHeaderFooterEvents = [event];
        } else if (tagName == 'drawing' || tagName.endsWith(':drawing')) {
          final rId = _getAttr(event, 'id');
          if (rId != null) {
            sheetObject._drawingRId = rId;
          }
        } else if (tagName == 'mergeCell' || tagName.endsWith(':mergeCell')) {
          final ref = _getAttr(event, 'ref');
          if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
            if (!sheetObject._spannedItems.contains(ref)) {
              sheetObject._spannedItems.add(ref);
            }
            final parts = ref.split(':');
            final startCell = parts[0];
            final endCell = parts[1];
            final spanObj = _Span.fromCellIndex(
              start: CellIndex.indexByString(startCell),
              end: CellIndex.indexByString(endCell),
            );
            if (!sheetObject._spanList.contains(spanObj)) {
              sheetObject._spanList.add(spanObj);
              // Clear merged cells from in-memory sheetData
              for (var col = spanObj.columnSpanStart; col <= spanObj.columnSpanEnd; col++) {
                for (var row = spanObj.rowSpanStart; row <= spanObj.rowSpanEnd; row++) {
                  final isOrigin = col == spanObj.columnSpanStart && row == spanObj.rowSpanStart;
                  if (!isOrigin) sheetObject._removeCell(row, col);
                }
              }
            }
            _excel._mergeChangeLookup = name;
          }
        } else if (currentCellEvents != null) {
          currentCellEvents.add(event);
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      } else if (event is xml_events.XmlEndElementEvent) {
        final tagName = event.name;

        if ((tagName == 'c' || tagName.endsWith(':c')) && currentCellEvents != null) {
          currentCellEvents.add(event);
          final cellXml = currentCellEvents.map((e) => e.toString()).join();
          final cellNode = XmlDocument.parse(cellXml).rootElement;
          if (currentWorksheetRowIndex != null && currentWorksheetRowIndex >= 0) {
            _parseCell(cellNode, sheetObject, currentWorksheetRowIndex, name);
          }
          currentCellEvents = null;
        } else if ((tagName == 'headerFooter' || tagName.endsWith(':headerFooter')) && currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
          final hfXml = currentHeaderFooterEvents.map((e) => e.toString()).join();
          final hfNode = XmlDocument.parse(hfXml).rootElement;
          sheetObject.headerFooter = HeaderFooter.fromXmlElement(hfNode);
          currentHeaderFooterEvents = null;
        } else if (tagName == 'row' || tagName.endsWith(':row')) {
          currentWorksheetRowIndex = null;
        } else if (currentCellEvents != null) {
          currentCellEvents.add(event);
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      } else {
        if (currentCellEvents != null) {
          currentCellEvents.add(event);
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      }
    }

    normalizeTable(sheetObject);
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  String? _getAttr(xml_events.XmlStartElementEvent event, String name) {
    for (final attr in event.attributes) {
      if (attr.name == name || attr.name.endsWith(':$name')) {
        return attr.value;
      }
    }
    return null;
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
}
