part of '../../../excel_community.dart';

class _WorksheetManager {
  final Excel _excel;
  final Save _save;
  final Map<CellStyle, int> _styleIndexCache = {};

  _WorksheetManager(this._excel, this._save);

  void setSheetElements() {
    _excel._sharedStrings.clear();
    _styleIndexCache.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (!_excel._xmlSheetId.containsKey(sheetName)) {
        _save.parser._createSheet(sheetName);
      }

      final sheetId = _excel._xmlSheetId[sheetName]!;
      final originalXmlString = _excel._sheetXmls[sheetId]!;

      final transformedXml =
          _transformWorksheetXml(sheetName, sheetObject, originalXmlString);
      _excel._sheetXmls[sheetId] = transformedXml;
    });
  }

  String _transformWorksheetXml(
      String sheetName, Sheet sheetObject, String originalXml) {
    final events = xml_events.parseEvents(originalXml);

    final worksheetAttributes = <xml_events.XmlEventAttribute>[];
    final originalElements = <String, List<String>>{};

    List<xml_events.XmlEvent>? currentCapturedEvents;
    String? currentTagName;
    int depth = 0;

    const replacedTags = {
      'sheetViews',
      'sheetFormatPr',
      'cols',
      'sheetData',
      'mergeCells',
      'headerFooter',
      'drawing',
    };

    for (final event in events) {
      if (event is xml_events.XmlStartElementEvent) {
        final tagName = event.name;
        if (tagName == 'worksheet') {
          worksheetAttributes.addAll(event.attributes);
          continue;
        }

        if (depth == 0) {
          currentTagName = tagName;
          currentCapturedEvents = [event];
          if (event.isSelfClosing) {
            final xmlString = event.toString();
            if (!replacedTags.contains(currentTagName)) {
              originalElements
                  .putIfAbsent(currentTagName, () => [])
                  .add(xmlString);
            }
            currentCapturedEvents = null;
            currentTagName = null;
          } else {
            depth = 1;
          }
        } else {
          currentCapturedEvents?.add(event);
          if (!event.isSelfClosing) {
            depth++;
          }
        }
      } else if (event is xml_events.XmlEndElementEvent) {
        final tagName = event.name;
        if (tagName == 'worksheet') {
          continue;
        }

        if (currentCapturedEvents != null) {
          currentCapturedEvents.add(event);
          depth--;
          if (depth == 0) {
            final xmlString =
                currentCapturedEvents.map((e) => e.toString()).join();
            if (!replacedTags.contains(currentTagName)) {
              originalElements
                  .putIfAbsent(currentTagName!, () => [])
                  .add(xmlString);
            }
            currentCapturedEvents = null;
            currentTagName = null;
          }
        }
      } else {
        if (currentCapturedEvents != null) {
          currentCapturedEvents.add(event);
        }
      }
    }

    final out = StringBuffer();
    out.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    out.write('<worksheet');
    for (final attr in worksheetAttributes) {
      out.write(' ${attr.name}="${attr.value}"');
    }
    out.write('>');

    final printedTags = <String>{};

    void writeOriginal(String tag) {
      printedTags.add(tag);
      final list = originalElements[tag];
      if (list != null) {
        for (final item in list) {
          out.write(item);
        }
      }
    }

    // Write in schema-compliant order:
    // 1. sheetPr
    writeOriginal('sheetPr');
    // 2. dimension
    writeOriginal('dimension');
    // 3. sheetViews
    out.write(_buildSheetViewsXml(sheetObject));
    printedTags.add('sheetViews');
    // 4. sheetFormatPr
    out.write(_buildSheetFormatPrXml(sheetObject));
    printedTags.add('sheetFormatPr');
    // 5. cols
    out.write(_buildColsXml(sheetObject));
    printedTags.add('cols');
    // 6. sheetData
    out.write(_buildSheetDataXml(sheetName, sheetObject));
    printedTags.add('sheetData');

    // Write common other elements
    writeOriginal('sheetProtection');
    writeOriginal('autoFilter');
    writeOriginal('sortState');
    writeOriginal('dataConsolidate');
    writeOriginal('customSheetViews');
    writeOriginal('customProperties');
    writeOriginal('cellWatches');
    writeOriginal('webPublishItems');
    writeOriginal('conditionalFormatting');
    writeOriginal('dataValidations');
    writeOriginal('hyperlinks');
    writeOriginal('printOptions');
    writeOriginal('pageMargins');
    writeOriginal('pageSetup');

    // 7. mergeCells
    out.write(_buildMergeCellsXml(sheetObject));
    printedTags.add('mergeCells');

    // 8. drawing / legacyDrawing / picture / oleObjects
    if (sheetObject._drawingRId != null) {
      out.write('<drawing r:id="${sheetObject._drawingRId}"/>');
    }
    printedTags.add('drawing');

    writeOriginal('legacyDrawing');
    writeOriginal('legacyDrawingHF');
    writeOriginal('picture');
    writeOriginal('oleObjects');
    writeOriginal('drawingHF');

    // 9. headerFooter
    out.write(_buildHeaderFooterXml(sheetObject));
    printedTags.add('headerFooter');

    // 10. extLst
    writeOriginal('extLst');

    // Catch-all: Write any other original elements we missed, just in case
    originalElements.forEach((tag, list) {
      if (!printedTags.contains(tag)) {
        for (final item in list) {
          out.write(item);
        }
      }
    });

    out.write('</worksheet>');
    return out.toString();
  }

  String _buildSheetViewsXml(Sheet sheetObject) {
    final buffer = StringBuffer();
    buffer.write('<sheetViews>');
    buffer.write('<sheetView workbookViewId="0"');
    if (sheetObject.isRTL) {
      buffer.write(' rightToLeft="1"');
    }
    buffer.write('/></sheetViews>');
    return buffer.toString();
  }

  String _buildSheetFormatPrXml(Sheet sheetObject) {
    final defaultRowHeight = sheetObject.defaultRowHeight;
    final defaultColumnWidth = sheetObject.defaultColumnWidth;

    if (defaultRowHeight == null && defaultColumnWidth == null) {
      return '';
    }
    final buffer = StringBuffer();
    buffer.write('<sheetFormatPr');
    if (defaultRowHeight != null) {
      buffer
          .write(' defaultRowHeight="${defaultRowHeight.toStringAsFixed(2)}"');
    }
    if (defaultColumnWidth != null) {
      buffer
          .write(' defaultColWidth="${defaultColumnWidth.toStringAsFixed(2)}"');
    }
    buffer.write('/>');
    return buffer.toString();
  }

  String _buildColsXml(Sheet sheetObject) {
    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    if (customWidths.isEmpty && autoFits.isEmpty) {
      return '';
    }

    final columnCount = max(
        autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
        customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1);

    final buffer = StringBuffer();
    buffer.write('<cols>');

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColumnWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }
      buffer.write(
          '<col min="${index + 1}" max="${index + 1}" width="${width.toStringAsFixed(2)}" bestFit="1" customWidth="1"/>');
    }
    buffer.write('</cols>');
    return buffer.toString();
  }

  String _buildSheetDataXml(String sheetName, Sheet sheetObject) {
    final customHeights = sheetObject.getRowHeights;
    final buffer = StringBuffer();
    buffer.write('<sheetData>');

    final sheetStyleReferenced = _excel._cellStyleReferenced[sheetName];

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      final rowData = sheetObject._sheetData[rowIndex];
      if (rowData == null || rowData.isEmpty) {
        continue;
      }

      double? height = customHeights[rowIndex];
      buffer.write('<row r="${rowIndex + 1}"');
      if (height != null) {
        buffer.write(' ht="${height.toStringAsFixed(2)}" customHeight="1"');
      }
      buffer.write('>');

      bool isSorted = true;
      int lastCol = -1;
      for (final colIndex in rowData.keys) {
        if (colIndex < lastCol) {
          isSorted = false;
          break;
        }
        lastCol = colIndex;
      }

      if (isSorted) {
        rowData.forEach((columnIndex, data) {
          _buildCellXml(
            buffer,
            sheetName,
            columnIndex,
            rowIndex,
            data.value,
            data._cellStyle,
            sheetStyleReferenced,
          );
        });
      } else {
        final sortedCols = rowData.keys.toList()..sort();
        for (final columnIndex in sortedCols) {
          final data = rowData[columnIndex]!;
          _buildCellXml(
            buffer,
            sheetName,
            columnIndex,
            rowIndex,
            data.value,
            data._cellStyle,
            sheetStyleReferenced,
          );
        }
      }
      buffer.write('</row>');
    }
    buffer.write('</sheetData>');
    return buffer.toString();
  }

  void _buildCellXml(
      StringBuffer buffer,
      String sheet,
      int columnIndex,
      int rowIndex,
      CellValue? value,
      CellStyle? cellStyle,
      Map<String, int>? sheetStyleReferenced) {
    SharedString? sharedString;
    if (value is TextCellValue) {
      final String strVal =
          (value.value.children == null || value.value.children!.isEmpty)
              ? (value.value.text ?? '')
              : value.toString();
      sharedString = _excel._sharedStrings.tryFind(strVal);
      if (sharedString != null) {
        _excel._sharedStrings.add(sharedString, strVal);
      } else {
        sharedString = _excel._sharedStrings.addFromString(strVal);
      }
    }

    String sAttr = '';
    if (_excel._styleChanges && cellStyle != null) {
      var upperLevelPos = _styleIndexCache[cellStyle];
      if (upperLevelPos == null) {
        upperLevelPos = _checkPosition(_excel._cellStyleList, cellStyle);
        if (upperLevelPos == -1) {
          int lowerLevelPos = _checkPosition(_save._innerCellStyle, cellStyle);
          if (lowerLevelPos != -1) {
            upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
          } else {
            upperLevelPos = 0;
          }
        }
        _styleIndexCache[cellStyle] = upperLevelPos;
      }
      sAttr = ' s="$upperLevelPos"';
    } else if (sheetStyleReferenced != null) {
      final rC = getCellId(columnIndex, rowIndex);
      if (sheetStyleReferenced.containsKey(rC)) {
        sAttr = ' s="${sheetStyleReferenced[rC]}"';
      }
    }

    String tAttr = '';
    if (value is TextCellValue) {
      tAttr = ' t="s"';
    } else if (value is BoolCellValue) {
      tAttr = ' t="b"';
    }

    buffer.write('<c r="');
    if (columnIndex < 16384) {
      buffer.write(_columnLetterList[columnIndex]);
    } else {
      buffer.write(_numericToLetters(columnIndex + 1));
    }
    buffer.write(rowIndex + 1);
    buffer.write('"');

    if (sAttr.isNotEmpty) {
      buffer.write(sAttr);
    }
    if (tAttr.isNotEmpty) {
      buffer.write(tAttr);
    }
    buffer.write('>');

    if (value != null) {
      final numFormat = cellStyle?.numberFormat;
      switch (value) {
        case TextCellValue():
          buffer.write('<v>');
          buffer.write(_excel._sharedStrings.indexOf(sharedString!));
          buffer.write('</v>');
          break;
        case FormulaCellValue():
          buffer.write('<f>');
          buffer.write(_escapeXml(value.formula));
          buffer.write('</f><v>');
          buffer.write(_escapeXml(value.write(numFormat)));
          buffer.write('</v>');
          break;
        case IntCellValue() || DoubleCellValue() || BoolCellValue():
          buffer.write('<v>');
          buffer.write(value.write(numFormat));
          buffer.write('</v>');
          break;
        case DateCellValue() || TimeCellValue() || DateTimeCellValue():
          buffer.write('<v>');
          buffer.write(value.write(numFormat));
          buffer.write('</v>');
          break;
      }
    }
    buffer.write('</c>');
  }

  String _buildMergeCellsXml(Sheet sheetObject) {
    final spannedItems = sheetObject.spannedItems;
    if (spannedItems.isEmpty) return '';

    final buffer = StringBuffer();
    buffer.write('<mergeCells count="${spannedItems.length}">');
    for (final ref in spannedItems) {
      buffer.write('<mergeCell ref="$ref"/>');
    }
    buffer.write('</mergeCells>');
    return buffer.toString();
  }

  String _buildHeaderFooterXml(Sheet sheetObject) {
    if (sheetObject.headerFooter == null) return '';
    return sheetObject.headerFooter!.toXmlElement().toXmlString();
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters =
            max(value[column]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  int _checkPosition(List<dynamic> list, dynamic element) {
    return list.indexOf(element);
  }
}
