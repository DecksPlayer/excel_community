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
    sheet._countRowsAndColumns();
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
  }

  /// Parses a single `<sheet>` node from the workbook, loads its XML file,
  /// parses it using SAX event streaming, and populates the [Sheet] object.
  void parseTable(XmlElement node) {
    final name = node.getAttribute('name')!;
    final target = _worksheetTargets[node.getAttribute('r:id')];
    if (target == null) {
      throw ArgumentError(
          'Worksheet target not found for relationship ID ${node.getAttribute('r:id')}');
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

    List<xml_events.XmlEvent>? currentHeaderFooterEvents;
    int? currentWorksheetRowIndex;

    // SAX cell parsing state
    bool insideCell = false;
    bool insideFormula = false;
    bool insideValue = false;
    bool insideInlineText = false;

    String? currentCellRef;
    String? currentCellType;
    String? currentCellStyleAttr;

    String? valueText;
    String? formulaText;
    String? inlineText;

    bool insideSheetView = false;

    for (final event in events) {
      if (event is xml_events.XmlStartElementEvent) {
        final tagName = event.name;

        if (tagName == 'sheetView' || tagName.endsWith(':sheetView')) {
          final rtl = _getAttr(event, 'rightToLeft');
          sheetObject.isRTL = rtl == '1';
          insideSheetView = true;
        } else if (insideSheetView &&
            (tagName == 'pane' || tagName.endsWith(':pane'))) {
          // <pane xSplit="..." ySplit="..." topLeftCell="..." state="frozen" />
          final xSplit = int.tryParse(_getAttr(event, 'xSplit') ?? '');
          final ySplit = int.tryParse(_getAttr(event, 'ySplit') ?? '');
          if ((xSplit ?? 0) > 0) {
            sheetObject._frozenColumns = xSplit;
          }
          if ((ySplit ?? 0) > 0) {
            sheetObject._frozenRows = ySplit;
          }
        } else if (tagName == 'sheetFormatPr' ||
            tagName.endsWith(':sheetFormatPr')) {
          final colW =
              double.tryParse(_getAttr(event, 'defaultColWidth') ?? '');
          final rowH =
              double.tryParse(_getAttr(event, 'defaultRowHeight') ?? '');
          if (colW != null && rowH != null) {
            sheetObject._defaultColumnWidth = colW;
            sheetObject._defaultRowHeight = rowH;
          }
        } else if (tagName == 'col' || tagName.endsWith(':col')) {
          final min = int.tryParse(_getAttr(event, 'min') ?? '');
          final maxVal = int.tryParse(_getAttr(event, 'max') ?? '');
          final width = double.tryParse(_getAttr(event, 'width') ?? '');
          final hiddenVal = _getAttr(event, 'hidden');
          final isHidden = hiddenVal == '1' || hiddenVal == 'true';
          if (min != null) {
            final end = maxVal ?? min;
            for (int col = min; col <= end; col++) {
              final zeroBasedCol = col - 1;
              if (zeroBasedCol >= 0) {
                if (width != null) {
                  sheetObject._columnWidths[zeroBasedCol] = width;
                }
                if (isHidden) {
                  sheetObject._hiddenColumns.add(zeroBasedCol);
                }
              }
            }
          }
        } else if (tagName == 'row' || tagName.endsWith(':row')) {
          final rowNum = int.tryParse(_getAttr(event, 'r') ?? '');
          if (rowNum != null) {
            currentWorksheetRowIndex = rowNum - 1;
            final height = double.tryParse(_getAttr(event, 'ht') ?? '');
            final hiddenVal = _getAttr(event, 'hidden');
            final isHidden = hiddenVal == '1' || hiddenVal == 'true';
            if (currentWorksheetRowIndex >= 0) {
              if (height != null) {
                sheetObject._rowHeights[currentWorksheetRowIndex] = height;
              }
              if (isHidden) {
                sheetObject._hiddenRows.add(currentWorksheetRowIndex);
              }
            }
          }
        } else if (tagName == 'c' || tagName.endsWith(':c')) {
          insideCell = true;
          currentCellRef = _getAttr(event, 'r');
          currentCellType = _getAttr(event, 't');
          currentCellStyleAttr = _getAttr(event, 's');
          valueText = null;
          formulaText = null;
          inlineText = null;

          if (event.isSelfClosing) {
            insideCell = false;
            if (currentCellRef != null &&
                currentWorksheetRowIndex != null &&
                currentWorksheetRowIndex >= 0) {
              _processCellInline(
                ref: currentCellRef,
                type: currentCellType,
                styleAttr: currentCellStyleAttr,
                valueStr: null,
                formulaStr: null,
                inlineStr: null,
                sheetObject: sheetObject,
                rowIndex: currentWorksheetRowIndex,
                sheetName: name,
              );
            }
          }
        } else if (insideCell) {
          if (tagName == 'v' || tagName.endsWith(':v')) {
            insideValue = true;
          } else if (tagName == 'f' || tagName.endsWith(':f')) {
            insideFormula = true;
          } else if (tagName == 't' || tagName.endsWith(':t')) {
            insideInlineText = true;
          }
        } else if (tagName == 'headerFooter' ||
            tagName.endsWith(':headerFooter')) {
          currentHeaderFooterEvents = [event];
        } else if (tagName == 'drawing' || tagName.endsWith(':drawing')) {
          final rId = _getAttr(event, 'id');
          if (rId != null) {
            sheetObject._drawingRId = rId;
          }
        } else if (tagName == 'legacyDrawing' || tagName.endsWith(':legacyDrawing')) {
          final rId = _getAttr(event, 'id');
          if (rId != null) {
            sheetObject._legacyDrawingRId = rId;
          }
        } else if (tagName == 'sheetProtection' ||
            tagName.endsWith(':sheetProtection')) {
          sheetObject.sheetProtection = SheetProtection(
            sheet: _getAttr(event, 'sheet') == '1' ||
                _getAttr(event, 'sheet') == 'true' ||
                _getAttr(event, 'sheet') == null,
            objects: _getAttr(event, 'objects') == '1' ||
                _getAttr(event, 'objects') == 'true',
            scenarios: _getAttr(event, 'scenarios') == '1' ||
                _getAttr(event, 'scenarios') == 'true',
            formatCells: _getAttr(event, 'formatCells') == '1' ||
                _getAttr(event, 'formatCells') == 'true',
            formatColumns: _getAttr(event, 'formatColumns') == '1' ||
                _getAttr(event, 'formatColumns') == 'true',
            formatRows: _getAttr(event, 'formatRows') == '1' ||
                _getAttr(event, 'formatRows') == 'true',
            insertColumns: _getAttr(event, 'insertColumns') == '1' ||
                _getAttr(event, 'insertColumns') == 'true',
            insertRows: _getAttr(event, 'insertRows') == '1' ||
                _getAttr(event, 'insertRows') == 'true',
            insertHyperlinks: _getAttr(event, 'insertHyperlinks') == '1' ||
                _getAttr(event, 'insertHyperlinks') == 'true',
            deleteColumns: _getAttr(event, 'deleteColumns') == '1' ||
                _getAttr(event, 'deleteColumns') == 'true',
            deleteRows: _getAttr(event, 'deleteRows') == '1' ||
                _getAttr(event, 'deleteRows') == 'true',
            selectLockedCells: _getAttr(event, 'selectLockedCells') == '1' ||
                _getAttr(event, 'selectLockedCells') == 'true' ||
                _getAttr(event, 'selectLockedCells') == null,
            selectUnlockedCells:
                _getAttr(event, 'selectUnlockedCells') == '1' ||
                    _getAttr(event, 'selectUnlockedCells') == 'true' ||
                    _getAttr(event, 'selectUnlockedCells') == null,
            sort: _getAttr(event, 'sort') == '1' ||
                _getAttr(event, 'sort') == 'true',
            autoFilter: _getAttr(event, 'autoFilter') == '1' ||
                _getAttr(event, 'autoFilter') == 'true',
            pivotTables: _getAttr(event, 'pivotTables') == '1' ||
                _getAttr(event, 'pivotTables') == 'true',
          );
          final pw = _getAttr(event, 'password');
          if (pw != null) {
            sheetObject.sheetProtection.password = pw;
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
              for (var col = spanObj.columnSpanStart;
                  col <= spanObj.columnSpanEnd;
                  col++) {
                for (var row = spanObj.rowSpanStart;
                    row <= spanObj.rowSpanEnd;
                    row++) {
                  final isOrigin = col == spanObj.columnSpanStart &&
                      row == spanObj.rowSpanStart;
                  if (!isOrigin) sheetObject._removeCell(row, col);
                }
              }
            }
            _excel._mergeChangeLookup = name;
          }
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      } else if (event is xml_events.XmlTextEvent) {
        if (insideCell) {
          if (insideValue) {
            valueText = (valueText ?? '') + event.value;
          } else if (insideFormula) {
            formulaText = (formulaText ?? '') + event.value;
          } else if (insideInlineText) {
            inlineText = (inlineText ?? '') + event.value;
          }
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      } else if (event is xml_events.XmlEndElementEvent) {
        final tagName = event.name;

        if (tagName == 'c' || tagName.endsWith(':c')) {
          if (currentCellRef != null &&
              currentWorksheetRowIndex != null &&
              currentWorksheetRowIndex >= 0) {
            _processCellInline(
              ref: currentCellRef,
              type: currentCellType,
              styleAttr: currentCellStyleAttr,
              valueStr: valueText,
              formulaStr: formulaText,
              inlineStr: inlineText,
              sheetObject: sheetObject,
              rowIndex: currentWorksheetRowIndex,
              sheetName: name,
            );
          }
          insideCell = false;
        } else if (insideCell) {
          if (tagName == 'v' || tagName.endsWith(':v')) {
            insideValue = false;
          } else if (tagName == 'f' || tagName.endsWith(':f')) {
            insideFormula = false;
          } else if (tagName == 't' || tagName.endsWith(':t')) {
            insideInlineText = false;
          }
        } else if (tagName == 'sheetView' || tagName.endsWith(':sheetView')) {
          insideSheetView = false;
        } else if ((tagName == 'headerFooter' ||
                tagName.endsWith(':headerFooter')) &&
            currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
          final hfXml =
              currentHeaderFooterEvents.map((e) => e.toString()).join();
          final hfNode = XmlDocument.parse(hfXml).rootElement;
          sheetObject.headerFooter = HeaderFooter.fromXmlElement(hfNode);
          currentHeaderFooterEvents = null;
        } else if (tagName == 'row' || tagName.endsWith(':row')) {
          currentWorksheetRowIndex = null;
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        }
      } else {
        if (currentHeaderFooterEvents != null) {
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

  void _processCellInline({
    required String ref,
    required String? type,
    required String? styleAttr,
    required String? valueStr,
    required String? formulaStr,
    required String? inlineStr,
    required Sheet sheetObject,
    required int rowIndex,
    required String sheetName,
  }) {
    final coords = _cellCoordsFromCellId(ref);
    final columnIndex = coords.$2;

    int s = 0;
    if (styleAttr != null) {
      try {
        s = int.parse(styleAttr);
      } catch (_) {}

      if (s > 0) {
        if (_excel._cellStyleReferenced[sheetName] == null) {
          _excel._cellStyleReferenced[sheetName] = {ref: s};
        } else {
          _excel._cellStyleReferenced[sheetName]![ref] = s;
        }
      }
    }

    CellValue? value;
    switch (type) {
      // Shared string
      case 's':
        if (valueStr != null) {
          final idx = int.tryParse(valueStr.trim()) ?? 0;
          final sharedStringVal = _excel._sharedStrings.value(idx);
          if (sharedStringVal != null) {
            value = TextCellValue.span(sharedStringVal.textSpan);
          }
        }
        break;
      // Boolean
      case 'b':
        value = BoolCellValue(valueStr == '1');
        break;
      // Error / formula string
      case 'e':
      case 'str':
        if (valueStr != null) {
          value = FormulaCellValue(valueStr);
        }
        break;
      // Inline string
      case 'inlineStr':
        if (inlineStr != null) {
          value = TextCellValue(inlineStr);
        }
        break;
      // Number (default)
      case 'n':
      default:
        if (formulaStr != null) {
          value = FormulaCellValue(formulaStr);
        } else if (valueStr != null) {
          if (styleAttr != null && s < _excel._cellStyleList.length) {
            final numFmtId = _excel._numFmtIds[s];
            final numFormat = _excel._numFormats.getByNumFmtId(numFmtId) ??
                NumFormat.standard_0;
            value = numFormat.read(valueStr);
          } else {
            value = NumFormat.defaultNumeric.read(valueStr);
          }
        }
    }

    final cell = Data.newData(sheetObject, rowIndex, columnIndex);
    cell._value = value;
    if (s < _excel._cellStyleList.length) {
      cell._cellStyle = _excel._cellStyleList[s];
    } else {
      cell._cellStyle = _excel._getDefaultStyle(NumFormat.defaultFor(value));
    }

    var rowData = sheetObject._sheetData[rowIndex];
    if (rowData == null) {
      sheetObject._sheetData[rowIndex] = rowData = {};
    }
    rowData[columnIndex] = cell;
  }
}
