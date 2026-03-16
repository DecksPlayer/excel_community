part of excel_community;

class Sheet {
  final Excel _excel;
  final String _sheet;
  bool _isRTL = false;
  int _maxRows = 0;
  int _maxColumns = 0;
  double? _defaultColumnWidth;
  double? _defaultRowHeight;
  Map<int, double> _columnWidths = {};
  Map<int, double> _rowHeights = {};
  Map<int, bool> _columnAutoFit = {};
  FastList<String> _spannedItems = FastList<String>();
  List<_Span?> _spanList = [];
  Map<int, Map<int, Data>> _sheetData = {};
  HeaderFooter? _headerFooter;
  final List<Chart> _charts = [];

  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
      : this._(excel, sheetName,
            sh: oldSheetObject._sheetData,
            spanL_: oldSheetObject._spanList,
            spanI_: oldSheetObject._spannedItems,
            maxRowsVal: oldSheetObject._maxRows,
            maxColumnsVal: oldSheetObject._maxColumns,
            columnWidthsVal: oldSheetObject._columnWidths,
            rowHeightsVal: oldSheetObject._rowHeights,
            columnAutoFitVal: oldSheetObject._columnAutoFit,
            isRTLVal: oldSheetObject._isRTL,
            headerFooter: oldSheetObject._headerFooter,
            charts: oldSheetObject._charts);

  Sheet._(this._excel, this._sheet,
      {Map<int, Map<int, Data>>? sh,
      List<_Span?>? spanL_,
      FastList<String>? spanI_,
      int? maxRowsVal,
      int? maxColumnsVal,
      bool? isRTLVal,
      Map<int, double>? columnWidthsVal,
      Map<int, double>? rowHeightsVal,
      Map<int, bool>? columnAutoFitVal,
      HeaderFooter? headerFooter,
      List<Chart>? charts}) {
    _headerFooter = headerFooter;
    if (charts != null) {
      _charts.addAll(charts);
    }

    if (spanL_ != null) {
      _spanList = List<_Span?>.from(spanL_);
      _excel._mergeChangeLookup = sheetName;
    }
    if (spanI_ != null) {
      _spannedItems = FastList<String>.from(spanI_);
    }
    if (maxColumnsVal != null) {
      _maxColumns = maxColumnsVal;
    }
    if (maxRowsVal != null) {
      _maxRows = maxRowsVal;
    }
    if (isRTLVal != null) {
      _isRTL = isRTLVal;
      _excel._rtlChangeLookup = sheetName;
    }
    if (columnWidthsVal != null) {
      _columnWidths = Map<int, double>.from(columnWidthsVal);
    }
    if (rowHeightsVal != null) {
      _rowHeights = Map<int, double>.from(rowHeightsVal);
    }
    if (columnAutoFitVal != null) {
      _columnAutoFit = Map<int, bool>.from(columnAutoFitVal);
    }

    if (sh != null) {
      _sheetData = <int, Map<int, Data>>{};
      sh.forEach((key, value) {
        _sheetData[key] = <int, Data>{};
        value.forEach((key1, oldDataObject) {
          _sheetData[key]![key1] = Data._clone(this, oldDataObject);
        });
      });
    }
    _countRowsAndColumns();
  }

  void _removeCell(int rowIndex, int columnIndex) {
    _sheetData[rowIndex]?.remove(columnIndex);
    if (_sheetData[rowIndex]?.isEmpty == true) {
      _sheetData.remove(rowIndex);
    }
  }

  bool get isRTL => _isRTL;

  set isRTL(bool _u) {
    _isRTL = _u;
    _excel._rtlChangeLookup = sheetName;
  }

  String get sheetName => _sheet;

  int get maxRows => _maxRows;

  int get maxColumns => _maxColumns;

  HeaderFooter? get headerFooter => _headerFooter;

  set headerFooter(HeaderFooter? headerFooter) {
    _headerFooter = headerFooter;
  }

  Data cell(CellIndex cellIndex) {
    _checkMaxColumn(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);

    if (cellIndex.columnIndex < 0 || cellIndex.rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex.columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex.columnIndex < 0 ? cellIndex.columnIndex : cellIndex.rowIndex} Negative index does not exist.');
    }

    if (_maxRows < (cellIndex.rowIndex + 1)) {
      _maxRows = cellIndex.rowIndex + 1;
    }

    if (_maxColumns < (cellIndex.columnIndex + 1)) {
      _maxColumns = cellIndex.columnIndex + 1;
    }

    if (_sheetData[cellIndex.rowIndex] != null) {
      if (_sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] == null) {
        _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] =
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex);
      }
    } else {
      _sheetData[cellIndex.rowIndex] = {
        cellIndex.columnIndex:
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex)
      };
    }

    return _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex]!;
  }

  void _putData(int rowIndex, int columnIndex, CellValue? value) {
    var row = _sheetData[rowIndex];
    if (row == null) {
      _sheetData[rowIndex] = row = {};
    }
    var cell = row[columnIndex];
    if (cell == null) {
      row[columnIndex] = cell = Data.newData(this, rowIndex, columnIndex);
    }

    cell._value = value;
    final currentStyle = cell._cellStyle;
    final defaultFormat = NumFormat.defaultFor(value);

    if (currentStyle == null) {
      cell._cellStyle = CellStyle(numberFormat: defaultFormat);
      if (defaultFormat != NumFormat.standard_0) {
        _excel._styleChanges = true;
      }
    } else {
      // Safer exhaustive switch as requested by USER
      final bool needsFormatUpdate;
      switch (value) {
        case null:
          needsFormatUpdate = currentStyle.numberFormat != NumFormat.standard_0;
        case FormulaCellValue() || TextCellValue():
          needsFormatUpdate = currentStyle.numberFormat == NumFormat.standard_0 &&
              defaultFormat != NumFormat.standard_0;
        case IntCellValue() || DoubleCellValue():
          needsFormatUpdate = !currentStyle.numberFormat.accepts(value) ||
              (currentStyle.numberFormat == NumFormat.standard_0 &&
                  defaultFormat != NumFormat.standard_0);
        case DateCellValue() || TimeCellValue() || DateTimeCellValue():
          // For temporal types, we're more aggressive about ensuring a temporal format
          needsFormatUpdate = !currentStyle.numberFormat.accepts(value) ||
              (currentStyle.numberFormat == NumFormat.standard_0 &&
                  defaultFormat != NumFormat.standard_0);
        case BoolCellValue():
          needsFormatUpdate = currentStyle.numberFormat == NumFormat.standard_0 &&
              defaultFormat != NumFormat.standard_0;
      }

      if (needsFormatUpdate) {
        cell._cellStyle = currentStyle.copyWith(
            numberFormat: (value == null) ? NumFormat.standard_0 : defaultFormat);
        _excel._styleChanges = true;
      }
    }

    if ((_maxColumns - 1) < columnIndex) {
      _maxColumns = columnIndex + 1;
    }

    if ((_maxRows - 1) < rowIndex) {
      _maxRows = rowIndex + 1;
    }
  }

  void _checkMaxColumn(int columnIndex) {
    if (_maxColumns >= 16384 || columnIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
    if (columnIndex < 0) {
      throw ArgumentError('Negative columnIndex found: $columnIndex');
    }
  }

  void _checkMaxRow(int rowIndex) {
    if (_maxRows >= 1048576 || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
    if (rowIndex < 0) {
      throw ArgumentError('Negative rowIndex found: $rowIndex');
    }
  }

  (int newRowIndex, int newColumnIndex) _isInsideSpanning(
      int rowIndex, int columnIndex) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex >= spanObj.rowSpanStart &&
          rowIndex <= spanObj.rowSpanEnd &&
          columnIndex >= spanObj.columnSpanStart &&
          columnIndex <= spanObj.columnSpanEnd) {
        newRowIndex = spanObj.rowSpanStart;
        newColumnIndex = spanObj.columnSpanStart;
        break;
      }
    }

    return (newRowIndex, newColumnIndex);
  }
}
