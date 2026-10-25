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

    final events = fastXmlEvents(contentString);

    List<xml_events.XmlEvent>? currentHeaderFooterEvents;
    List<xml_events.XmlEvent>? currentAutoFilterEvents;
    String? autoFilterRef;
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
    bool? fitToPage;
    Map<String, String>? relationshipTargets;
    // <dataValidation> being read: attributes and formula text.
    Map<String, String>? validationAttrs;
    final validationFormulas = <String, StringBuffer>{};
    String? validationFormula;

    for (final event in events) {
      if (event is xml_events.XmlStartElementEvent) {
        final tagName = event.name;

        if (tagName == 'tabColor' || tagName.endsWith(':tabColor')) {
          final rgb = _getAttr(event, 'rgb');
          final themeStr = _getAttr(event, 'theme');
          final tintStr = _getAttr(event, 'tint');
          final indexedStr = _getAttr(event, 'indexed');
          final autoStr = _getAttr(event, 'auto');

          final theme = themeStr != null ? int.tryParse(themeStr) : null;
          final tint = tintStr != null ? double.tryParse(tintStr) : null;
          final indexed = indexedStr != null ? int.tryParse(indexedStr) : null;
          final auto =
              autoStr != null ? (autoStr == '1' || autoStr == 'true') : null;

          sheetObject.tabColor = TabColor(
            rgb: rgb != null ? _normalizeColorHex(rgb) : null,
            color: rgb != null ? ExcelColor.fromHexString(rgb) : null,
            theme: theme,
            tint: tint,
            indexed: indexed,
            auto: auto,
          );
        } else if (tagName == 'outlinePr' || tagName.endsWith(':outlinePr')) {
          bool flag(String name, bool fallback) {
            final value = _getAttr(event, name);
            return value == null ? fallback : (value == '1' || value == 'true');
          }

          sheetObject.outlineSettings = OutlineSettings(
            summaryBelow: flag('summaryBelow', true),
            summaryRight: flag('summaryRight', true),
            showOutlineSymbols: flag('showOutlineSymbols', true),
            applyStyles: flag('applyStyles', false),
          );
        } else if (tagName == 'pageSetUpPr' ||
            tagName.endsWith(':pageSetUpPr')) {
          fitToPage = _parseBoolAttr(event, 'fitToPage');
        } else if (tagName == 'pageMargins' ||
            tagName.endsWith(':pageMargins')) {
          sheetObject._pageMargins = _parsePageMargins(event);
        } else if (tagName == 'printOptions' ||
            tagName.endsWith(':printOptions')) {
          sheetObject.printOptions = PrintOptions(
            gridLines: _parseBoolAttr(event, 'gridLines'),
            headings: _parseBoolAttr(event, 'headings'),
            horizontalCentered: _parseBoolAttr(event, 'horizontalCentered'),
            verticalCentered: _parseBoolAttr(event, 'verticalCentered'),
            gridLinesSet: _parseBoolAttr(event, 'gridLinesSet'),
          );
        } else if ((tagName == 'dataValidation' ||
                tagName.endsWith(':dataValidation')) &&
            !tagName.startsWith('x14:')) {
          // Excel 2010 x14:dataValidation (in extLst) is kept as raw XML.
          validationAttrs = {
            for (final attr in event.attributes)
              attr.name.split(':').last: attr.value,
          };
          validationFormulas.clear();
          if (event.isSelfClosing) {
            _addParsedValidation(sheetObject, validationAttrs, validationFormulas);
            validationAttrs = null;
          }
        } else if (validationAttrs != null &&
            (tagName == 'formula1' || tagName == 'formula2')) {
          validationFormula = tagName;
          validationFormulas[tagName] = StringBuffer();
        } else if (tagName == 'tablePart' || tagName.endsWith(':tablePart')) {
          final rId = _getAttr(event, 'id');
          if (rId != null) {
            relationshipTargets ??= _relationshipTargets(path);
            final target = relationshipTargets[rId];
            if (target != null) _addParsedTable(sheetObject, path, target);
          }
        } else if (tagName == 'hyperlink' || tagName.endsWith(':hyperlink')) {
          final ref = _getAttr(event, 'ref');
          final rId = _getAttr(event, 'id');
          String? url;
          if (rId != null) {
            relationshipTargets ??= _relationshipTargets(path);
            url = relationshipTargets[rId];
          }
          final location = _getAttr(event, 'location');
          if (ref != null && (url != null || location != null)) {
            try {
              sheetObject._hyperlinks[_CellRect.parse(ref).ref] = Hyperlink(
                url: url,
                location: location,
                tooltip: _getAttr(event, 'tooltip'),
                display: _getAttr(event, 'display'),
              );
            } catch (_) {
              // Ignore malformed references.
            }
          }
        } else if (tagName == 'pageSetup' || tagName.endsWith(':pageSetup')) {
          sheetObject.pageSetup = _parsePageSetup(event);
        } else if (tagName == 'sheetView' || tagName.endsWith(':sheetView')) {
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
          final outlineLevel = int.tryParse(_getAttr(event, 'outlineLevel') ?? '') ?? 0;
          final isCollapsed = _parseBoolAttr(event, 'collapsed') ?? false;
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
                if (outlineLevel > 0) {
                  sheetObject._columnOutlineLevels[zeroBasedCol] =
                      outlineLevel.clamp(1, _maxOutlineLevel);
                }
                if (isCollapsed) {
                  sheetObject._collapsedColumns.add(zeroBasedCol);
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
              final outlineLevel =
                  int.tryParse(_getAttr(event, 'outlineLevel') ?? '') ?? 0;
              if (outlineLevel > 0) {
                sheetObject._rowOutlineLevels[currentWorksheetRowIndex] =
                    outlineLevel.clamp(1, _maxOutlineLevel);
              }
              if (_parseBoolAttr(event, 'collapsed') ?? false) {
                sheetObject._collapsedRows.add(currentWorksheetRowIndex);
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
        } else if (tagName == 'legacyDrawing' ||
            tagName.endsWith(':legacyDrawing')) {
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
        } else if (tagName == 'autoFilter' ||
            tagName.endsWith(':autoFilter')) {
          final ref = _getAttr(event, 'ref');
          if (event.isSelfClosing) {
            if (ref != null && ref.isNotEmpty) {
              sheetObject.autoFilter = AutoFilter(ref: ref);
            }
          } else {
            autoFilterRef = ref;
            currentAutoFilterEvents = [event];
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
              // Clear merged cells from in-memory sheetData (keeps origin cell).
              // Non-origin cell style references are preserved in
              // _cellStyleReferenced so the writer can emit style-only <c>
              // elements for them (needed to preserve borders, etc.).
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
        } else if (currentAutoFilterEvents != null) {
          currentAutoFilterEvents.add(event);
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
        } else if (validationFormula != null) {
          validationFormulas[validationFormula]!.write(event.value);
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        } else if (currentAutoFilterEvents != null) {
          currentAutoFilterEvents.add(event);
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
        } else if (validationFormula != null &&
            (tagName == 'formula1' || tagName == 'formula2')) {
          validationFormula = null;
        } else if (validationAttrs != null &&
            (tagName == 'dataValidation' || tagName.endsWith(':dataValidation'))) {
          _addParsedValidation(sheetObject, validationAttrs, validationFormulas);
          validationAttrs = null;
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
        } else if ((tagName == 'autoFilter' ||
                tagName.endsWith(':autoFilter')) &&
            currentAutoFilterEvents != null) {
          currentAutoFilterEvents.add(event);
          final afXml =
              currentAutoFilterEvents.map((e) => e.toString()).join();
          sheetObject.autoFilter = _parseAutoFilterXml(autoFilterRef, afXml);
          currentAutoFilterEvents = null;
          autoFilterRef = null;
        } else if (tagName == 'row' || tagName.endsWith(':row')) {
          currentWorksheetRowIndex = null;
        } else if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        } else if (currentAutoFilterEvents != null) {
          currentAutoFilterEvents.add(event);
        }
      } else {
        if (currentHeaderFooterEvents != null) {
          currentHeaderFooterEvents.add(event);
        } else if (currentAutoFilterEvents != null) {
          currentAutoFilterEvents.add(event);
        }
      }
    }

    if (fitToPage != null) {
      sheetObject.pageSetup = (sheetObject.pageSetup ?? const PageSetup())
          .copyWith(fitToPage: fitToPage);
    }

    _parseConditionalFormatting(sheetObject, contentString);
    normalizeTable(sheetObject);
  }

  void _addParsedValidation(Sheet sheetObject, Map<String, String> attrs,
      Map<String, StringBuffer> formulas) {
    final sqref = attrs['sqref'];
    if (sqref == null || sqref.trim().isEmpty) return;
    bool flag(String name) => attrs[name] == '1' || attrs[name] == 'true';
    final validation = DataValidation(
      type: DataValidationType.fromXmlValue(attrs['type']),
      operator: DataValidationOperator.fromXmlValue(attrs['operator']),
      formula1: formulas['formula1']?.toString(),
      formula2: formulas['formula2']?.toString(),
      allowBlank: flag('allowBlank'),
      showDropdown: !flag('showDropDown'),
      showInputMessage: flag('showInputMessage'),
      showErrorMessage: flag('showErrorMessage'),
      promptTitle: attrs['promptTitle'],
      prompt: attrs['prompt'],
      errorTitle: attrs['errorTitle'],
      error: attrs['error'],
      errorStyle: DataValidationErrorStyle.fromXmlValue(attrs['errorStyle']),
    );
    try {
      final key = _CellRect.parseList(sqref).map((r) => r.ref).join(' ');
      sheetObject._dataValidations[key] = validation;
    } catch (_) {
      // Ignore malformed ranges.
    }
  }

  /// Reads the table part [target] (relative to the worksheet) into
  /// [sheetObject].
  void _addParsedTable(Sheet sheetObject, String worksheetPath, String target) {
    var partPath = target;
    if (partPath.startsWith('/')) {
      partPath = partPath.substring(1);
    } else {
      final segments = worksheetPath.split('/')..removeLast();
      for (final part in target.split('/')) {
        if (part == '..') {
          if (segments.isNotEmpty) segments.removeLast();
        } else if (part != '.') {
          segments.add(part);
        }
      }
      partPath = segments.join('/');
    }
    final file = _excel._archive.findFile(partPath);
    if (file == null) return;
    try {
      file.decompress();
      final table = ExcelTable._fromXml(XmlDocument.parse(utf8.decode(file.content)));
      if (table != null) sheetObject._tables.add(table);
    } catch (_) {
      // Ignore unreadable table parts.
    }
  }

  /// Targets of the worksheet's relationships (`.rels`), keyed by id.
  Map<String, String> _relationshipTargets(String worksheetPath) {
    final slash = worksheetPath.lastIndexOf('/');
    final relsPath = '${worksheetPath.substring(0, slash)}/_rels/'
        '${worksheetPath.substring(slash + 1)}.rels';
    final file = _excel._archive.findFile(relsPath);
    if (file == null) return const {};
    file.decompress();
    final document = XmlDocument.parse(utf8.decode(file.content));
    return {
      for (final rel in document.findAllElements('Relationship'))
        if (rel.getAttribute('Id') != null && rel.getAttribute('Target') != null)
          rel.getAttribute('Id')!: rel.getAttribute('Target')!,
    };
  }

  bool? _parseBoolAttr(xml_events.XmlStartElementEvent event, String name) {
    final value = _getAttr(event, name);
    if (value == null) return null;
    return value == '1' || value.toLowerCase() == 'true';
  }

  int? _parseIntAttr(xml_events.XmlStartElementEvent event, String name) {
    final value = _getAttr(event, name);
    return value != null ? int.tryParse(value) : null;
  }

  PageMargins _parsePageMargins(xml_events.XmlStartElementEvent event) {
    double attr(String name, double fallback) =>
        double.tryParse(_getAttr(event, name) ?? '') ?? fallback;
    const d = PageMargins.normal;
    return PageMargins(
      left: attr('left', d.left),
      right: attr('right', d.right),
      top: attr('top', d.top),
      bottom: attr('bottom', d.bottom),
      header: attr('header', d.header),
      footer: attr('footer', d.footer),
    );
  }

  PageSetup? _parsePageSetup(xml_events.XmlStartElementEvent event) {
    final paperSizeCode = _parseIntAttr(event, 'paperSize');
    final setup = PageSetup(
      orientation: PageOrientation.fromXmlValue(_getAttr(event, 'orientation')),
      paperSize:
          paperSizeCode != null ? PaperSize.fromCode(paperSizeCode) : null,
      paperWidth: _getAttr(event, 'paperWidth'),
      paperHeight: _getAttr(event, 'paperHeight'),
      scale: _parseIntAttr(event, 'scale'),
      fitToWidth: _parseIntAttr(event, 'fitToWidth'),
      fitToHeight: _parseIntAttr(event, 'fitToHeight'),
      firstPageNumber: _parseIntAttr(event, 'firstPageNumber'),
      useFirstPageNumber: _parseBoolAttr(event, 'useFirstPageNumber'),
      pageOrder: PageOrder.fromXmlValue(_getAttr(event, 'pageOrder')),
      blackAndWhite: _parseBoolAttr(event, 'blackAndWhite'),
      draft: _parseBoolAttr(event, 'draft'),
      cellComments:
          PrintCellComments.fromXmlValue(_getAttr(event, 'cellComments')),
      errors: PrintErrors.fromXmlValue(_getAttr(event, 'errors')),
      horizontalDpi: _parseIntAttr(event, 'horizontalDpi'),
      verticalDpi: _parseIntAttr(event, 'verticalDpi'),
      copies: _parseIntAttr(event, 'copies'),
      usePrinterDefaults: _parseBoolAttr(event, 'usePrinterDefaults'),
    );
    return setup.hasPageSetupAttributes ? setup : null;
  }

  void _parseConditionalFormatting(Sheet sheetObject, String contentString) {
    if (!contentString.contains('conditionalFormatting')) return;
    try {
      final doc = XmlDocument.parse(contentString);
      final cfElements = doc.findAllElements('conditionalFormatting');
      for (final cf in cfElements) {
        final sqref = cf.getAttribute('sqref');
        if (sqref == null || sqref.isEmpty) continue;

        final rules = <ConditionalFormattingRule>[];
        for (final ruleNode in cf.findElements('cfRule')) {
          final typeStr = ruleNode.getAttribute('type') ?? 'cellIs';
          final type = ConditionalFormattingType.fromValue(typeStr);

          final opStr = ruleNode.getAttribute('operator');
          final operator = ConditionalFormattingOperator.fromValue(opStr);

          final priority =
              int.tryParse(ruleNode.getAttribute('priority') ?? '') ?? 1;
          final text = ruleNode.getAttribute('text');

          final dxfIdStr = ruleNode.getAttribute('dxfId');
          DifferentialStyle style = const DifferentialStyle();
          if (dxfIdStr != null) {
            final dxfId = int.tryParse(dxfIdStr);
            if (dxfId != null && dxfId >= 0 && dxfId < _excel._dxfList.length) {
              style = _excel._dxfList[dxfId];
            }
          }

          final formulae =
              ruleNode.findElements('formula').map((e) => e.innerText).toList();

          rules.add(ConditionalFormattingRule(
            type: type,
            operator: operator,
            formulae: formulae,
            style: style,
            priority: priority,
            text: text,
          ));
        }

        if (rules.isNotEmpty) {
          sheetObject._conditionalFormattings
              .add(ConditionalFormattingGroup(sqref: sqref, rules: rules));
        }
      }
    } catch (_) {}
  }

  AutoFilter _parseAutoFilterXml(String? refAttr, String xmlString) {
    try {
      final doc = XmlDocument.parse(xmlString);
      final root = doc.rootElement;
      final ref = refAttr ?? root.getAttribute('ref') ?? '';
      final filterCols = <FilterColumn>[];

      for (final fc in root.findAllElements('filterColumn')) {
        final colId = int.tryParse(fc.getAttribute('colId') ?? '0') ?? 0;
        final hb = fc.getAttribute('hiddenButton');
        final sb = fc.getAttribute('showButton');
        final hiddenButton = hb == null ? null : (hb == '1' || hb == 'true');
        final showButton = sb == null ? null : (sb == '1' || sb == 'true');

        final filterValues = <String>[];
        bool blank = false;
        final filtersElem = fc.findElements('filters').firstOrNull;
        if (filtersElem != null) {
          blank = filtersElem.getAttribute('blank') == '1' ||
              filtersElem.getAttribute('blank') == 'true';
          for (final f in filtersElem.findElements('filter')) {
            final val = f.getAttribute('val');
            if (val != null) {
              filterValues.add(val);
            }
          }
        }

        final customFilters = <CustomFilterRule>[];
        bool customFiltersAnd = false;
        final cfElem = fc.findElements('customFilters').firstOrNull;
        if (cfElem != null) {
          customFiltersAnd = cfElem.getAttribute('and') == '1' ||
              cfElem.getAttribute('and') == 'true';
          for (final cf in cfElem.findElements('customFilter')) {
            final opStr = cf.getAttribute('operator') ?? 'equal';
            final val = cf.getAttribute('val') ?? '';
            customFilters.add(CustomFilterRule(
              operator: FilterOperator.fromValue(opStr),
              val: val,
            ));
          }
        }

        bool hasComplex = false;
        for (final child in fc.childElements) {
          final localName = child.name.local;
          if (localName != 'filters' && localName != 'customFilters') {
            hasComplex = true;
            break;
          }
          if (localName == 'filters') {
            for (final sub in child.childElements) {
              if (sub.name.local != 'filter') {
                hasComplex = true;
                break;
              }
            }
          }
        }

        final innerChildren =
            hasComplex ? fc.children.map((c) => c.toXmlString()).join() : null;

        filterCols.add(FilterColumn(
          colId: colId,
          hiddenButton: hiddenButton,
          showButton: showButton,
          filterValues: filterValues,
          blank: blank,
          customFilters: customFilters,
          customFiltersAnd: customFiltersAnd,
          customXml: innerChildren,
        ));
      }

      final innerXml = root.children.map((c) => c.toXmlString()).join();
      return AutoFilter(
        ref: ref,
        filterColumns: filterCols,
        customXml: filterCols.isEmpty && innerXml.isNotEmpty ? innerXml : null,
      );
    } catch (_) {
      return AutoFilter(ref: refAttr ?? '');
    }
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
          CellValue? cachedValue;
          if (valueStr != null) {
            if (styleAttr != null && s < _excel._cellStyleList.length) {
              final numFmtId = _excel._numFmtIds[s];
              final numFormat = _excel._numFormats.getByNumFmtId(numFmtId) ??
                  NumFormat.standard_0;
              cachedValue = numFormat.read(valueStr);
            } else {
              cachedValue = NumFormat.defaultNumeric.read(valueStr);
            }
          }
          value = FormulaCellValue(formulaStr, cachedValue: cachedValue);
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
