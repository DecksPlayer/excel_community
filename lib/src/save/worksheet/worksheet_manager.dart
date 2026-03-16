part of '../../../excel_community.dart';

class _WorksheetManager {
  final Excel _excel;
  final Save _save;

  _WorksheetManager(this._excel, this._save);

  void setSheetElements() {
    _excel._sharedStrings.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (_excel._sheets[sheetName] == null) {
        _save.parser._createSheet(sheetName);
      }

      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      XmlElement worksheetElement = xmlFile.findAllElements('worksheet').first;

      XmlElement? sheetFormatPrElement =
          worksheetElement.findElements('sheetFormatPr').isNotEmpty
              ? worksheetElement.findElements('sheetFormatPr').first
              : null;

      if (sheetFormatPrElement != null) {
        sheetFormatPrElement.attributes.clear();

        if (defaultRowHeight == null && defaultColumnWidth == null) {
          worksheetElement.children.remove(sheetFormatPrElement);
        }
      } else if (defaultRowHeight != null || defaultColumnWidth != null) {
        sheetFormatPrElement = XmlElement(XmlName('sheetFormatPr'), [], []);
        worksheetElement.children.insert(0, sheetFormatPrElement);
      }

      if (defaultRowHeight != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultRowHeight'), defaultRowHeight.toStringAsFixed(2)));
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultColWidth'), defaultColumnWidth.toStringAsFixed(2)));
      }

      _setColumns(sheetObject, xmlFile);
      _setRows(sheetName, sheetObject);
      _save._setHeaderFooter(sheetName);
    });
  }

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final columnElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColumnWidths.isEmpty &&
        sheetObject.getColumnAutoFits.isEmpty) {
      if (columnElements.isEmpty) {
        return;
      }

      final columns = columnElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(columns);
      return;
    }

    if (columnElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(XmlName('cols'), [], []));
    }

    var columns = columnElements.first;

    if (columns.children.isNotEmpty) {
      columns.children.clear();
    }

    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    final columnCount = max(
        autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
        customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1);

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

      _addNewColumn(columns, index, index, width);
    }
  }

  void _setRows(String sheetName, Sheet sheetObject) {
    final customHeights = sheetObject.getRowHeights;

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      double? height;

      if (customHeights.containsKey(rowIndex)) {
        height = customHeights[rowIndex];
      }

      if (sheetObject._sheetData[rowIndex] == null) {
        continue;
      }
      var foundRow = _createNewRow(
          _excel._sheets[sheetName]! as XmlElement, rowIndex, height);
      for (var columnIndex = 0;
          columnIndex < sheetObject._maxColumns;
          columnIndex++) {
        var data = sheetObject._sheetData[rowIndex]![columnIndex];
        if (data == null) {
          continue;
        }
        _updateCell(sheetName, foundRow, columnIndex, rowIndex, data.value,
            data.cellStyle?.numberFormat);
      }
    }
  }

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
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

  XmlElement _createNewRow(XmlElement table, int rowIndex, double? height) {
    var row = XmlElement(XmlName('row'), [
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
      if (height != null)
        XmlAttribute(XmlName('ht'), height.toStringAsFixed(2)),
      if (height != null) XmlAttribute(XmlName('customHeight'), '1'),
    ], []);
    table.children.add(row);
    return row;
  }

  XmlElement _updateCell(String sheet, XmlElement row, int columnIndex,
      int rowIndex, CellValue? value, NumFormat? numberFormat) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value, numberFormat);
    row.children.add(cell);
    return cell;
  }

  XmlElement _createCell(String sheet, int columnIndex, int rowIndex,
      CellValue? value, NumFormat? numberFormat) {
    SharedString? sharedString;
    if (value is TextCellValue) {
      sharedString = _excel._sharedStrings.tryFind(value.toString());
      if (sharedString != null) {
        _excel._sharedStrings.add(sharedString, value.toString());
      } else {
        sharedString = _excel._sharedStrings.addFromString(value.toString());
      }
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      if (value is TextCellValue) XmlAttribute(XmlName('t'), 's'),
      if (value is BoolCellValue) XmlAttribute(XmlName('t'), 'b'),
    ];

    final cellStyle =
        _excel._sheetMap[sheet]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;

    if (_excel._styleChanges && cellStyle != null) {
      int upperLevelPos = _checkPosition(_excel._cellStyleList, cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _checkPosition(_save._innerCellStyle, cellStyle);
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '$upperLevelPos'),
      );
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._cellStyleReferenced[sheet]![rC]}'),
      );
    }


    final List<XmlElement> children;
    switch (value) {
      case null:
        children = [];
      case TextCellValue():
        children = [
          XmlElement(XmlName('v'), [], [
            XmlText(_excel._sharedStrings.indexOf(sharedString!).toString())
          ]),
        ];
      case FormulaCellValue():
        children = [
          XmlElement(XmlName('f'), [], [XmlText(value.formula)]),
          XmlElement(XmlName('v'), [], [XmlText(value.write(numberFormat))]),
        ];
      case IntCellValue() ||
            DoubleCellValue() ||
            DateCellValue() ||
            TimeCellValue() ||
            DateTimeCellValue() ||
            BoolCellValue():
        children = [
          XmlElement(XmlName('v'), [], [XmlText(value.write(numberFormat))]),
        ];
    }

    return XmlElement(XmlName('c'), attributes, children);
  }
}
