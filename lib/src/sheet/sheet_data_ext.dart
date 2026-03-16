part of excel_community;

extension SheetDataExt on Sheet {
  ///
  /// returns `2-D dynamic List` of the sheet elements
  ///
  List<List<Data?>> get rows {
    var _data = <List<Data?>>[];

    if (_sheetData.isEmpty) {
      return _data;
    }

    if (_maxRows > 0 && maxColumns > 0) {
      _data = List.generate(_maxRows, (rowIndex) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex] != null &&
              _sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex];
          }
          return null;
        });
      });
    }

    return _data;
  }

  /// returns row at index = `rowIndex`
  List<Data?> row(int rowIndex) {
    if (rowIndex < 0) {
      return <Data?>[];
    }
    if (rowIndex < _maxRows) {
      if (_sheetData[rowIndex] != null) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex]!;
          }
          return null;
        });
      } else {
        return List.generate(_maxColumns, (_) => null);
      }
    }
    return <Data?>[];
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<Data?>?> selectRangeWithString(String range) {
    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRange(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRange(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  List<List<Data?>?> selectRange(CellIndex start, {CellIndex? end}) {
    _checkMaxColumn(start.columnIndex);
    _checkMaxRow(start.rowIndex);
    if (end != null) {
      _checkMaxColumn(end.columnIndex);
      _checkMaxRow(end.rowIndex);
    }

    int _startColumn = start.columnIndex, _startRow = start.rowIndex;
    int? _endColumn = end?.columnIndex, _endRow = end?.rowIndex;

    if (_endColumn != null && _endRow != null) {
      if (_startRow > _endRow) {
        _startRow = end!.rowIndex;
        _endRow = start.rowIndex;
      }
      if (_endColumn < _startColumn) {
        _endColumn = start.columnIndex;
        _startColumn = end!.columnIndex;
      }
    }

    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (_sheetData.isEmpty) {
      return _selectedRange;
    }

    for (var i = _startRow; i <= (_endRow ?? maxRows); i++) {
      var mapData = _sheetData[i];
      if (mapData != null) {
        List<Data?> row = <Data?>[];
        for (var j = _startColumn; j <= (_endColumn ?? maxColumns); j++) {
          row.add(mapData[j]);
        }
        _selectedRange.add(row);
      } else {
        _selectedRange.add(null);
      }
    }

    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<dynamic>?> selectRangeValuesWithString(String range) {
    List<List<dynamic>?> _selectedRange = <List<dynamic>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRangeValues(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRangeValues(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  List<List<dynamic>?> selectRangeValues(CellIndex start, {CellIndex? end}) {
    var _list =
        (end == null ? selectRange(start) : selectRange(start, end: end));
    return _list
        .map((List<Data?>? e) =>
            e?.map((e1) => e1 != null ? e1.value : null).toList())
        .toList();
  }

  ///
  /// updates count of rows and columns
  ///
  _countRowsAndColumns() {
    int maximumColumnIndex = -1, maximumRowIndex = -1;
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (_sheetData[rowKey] != null && _sheetData[rowKey]!.isNotEmpty) {
        List<int> keys = _sheetData[rowKey]!.keys.toList()..sort();
        if (keys.isNotEmpty && keys.last > maximumColumnIndex) {
          maximumColumnIndex = keys.last;
        }
      }
    });

    if (sortedKeys.isNotEmpty) {
      maximumRowIndex = sortedKeys.last;
    }

    _maxColumns = maximumColumnIndex + 1;
    _maxRows = maximumRowIndex + 1;
  }

  ///
  /// If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0 || columnIndex >= maxColumns) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (columnIndex < spanObj.columnSpanStart) {
        // Shifting required
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart,
          columnSpanStart: spanObj.columnSpanStart - 1,
          rowSpanEnd: spanObj.rowSpanEnd,
          columnSpanEnd: spanObj.columnSpanEnd - 1,
        );
        updateSpanCell = true;
      } else if (columnIndex <= spanObj.columnSpanEnd) {
        // Shrink the span as the column is removed.
        if (spanObj.columnSpanStart == spanObj.columnSpanEnd) {
          // Both are same then it means that it will no longer span.
          _spanList[i] = null;
          updateSpanCell = true;
        } else {
          _spanList[i] = _Span(
            rowSpanStart: spanObj.rowSpanStart,
            columnSpanStart: spanObj.columnSpanStart,
            rowSpanEnd: spanObj.rowSpanEnd,
            columnSpanEnd: spanObj.columnSpanEnd - 1,
          );
          updateSpanCell = true;
        }
      }
    }

    if (updateSpanCell) {
      _cleanUpSpanMap();
    }

    for (int i = 0; i < maxRows; i++) {
      if (_sheetData[i] != null && _sheetData[i]!.containsKey(columnIndex)) {
        _sheetData[i]!.remove(columnIndex);
      }
      if (_sheetData[i] != null) {
        Map<int, Data> _map = <int, Data>{};
        List<int> sortedKeys = _sheetData[i]!.keys.toList()..sort();
        sortedKeys.forEach((columnKey) {
          if (columnKey > columnIndex) {
            _map[columnKey - 1] = _sheetData[i]![columnKey]!;
            _map[columnKey - 1]!._columnIndex = columnKey - 1;
          } else {
            _map[columnKey] = _sheetData[i]![columnKey]!;
          }
        });
        _sheetData[i] = _map;
      }
    }
    _countRowsAndColumns();
  }

  ///
  /// Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  void insertColumn(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (columnIndex <= spanObj.columnSpanStart) {
        // Shifting required
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart,
          columnSpanStart: spanObj.columnSpanStart + 1,
          rowSpanEnd: spanObj.rowSpanEnd,
          columnSpanEnd: spanObj.columnSpanEnd + 1,
        );
        updateSpanCell = true;
      } else if (columnIndex <= spanObj.columnSpanEnd) {
        // Expand the span as the column is inserted.
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart,
          columnSpanStart: spanObj.columnSpanStart,
          rowSpanEnd: spanObj.rowSpanEnd,
          columnSpanEnd: spanObj.columnSpanEnd + 1,
        );
        updateSpanCell = true;
      }
    }

    if (updateSpanCell) {
      _cleanUpSpanMap();
    }

    for (int i = 0; i < maxRows; i++) {
      if (_sheetData[i] != null) {
        Map<int, Data> _map = <int, Data>{};
        List<int> sortedKeys = _sheetData[i]!.keys.toList()..sort();
        sortedKeys.forEach((columnKey) {
          if (columnKey >= columnIndex) {
            _map[columnKey + 1] = _sheetData[i]![columnKey]!;
            _map[columnKey + 1]!._columnIndex = columnKey + 1;
          } else {
            _map[columnKey] = _sheetData[i]![columnKey]!;
          }
        });
        _sheetData[i] = _map;
      }
    }
    _countRowsAndColumns();
  }

  ///
  /// If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(int rowIndex) {
    _checkMaxRow(rowIndex);
    if (rowIndex < 0 || rowIndex >= maxRows) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex < spanObj.rowSpanStart) {
        // Shifting required
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart - 1,
          columnSpanStart: spanObj.columnSpanStart,
          rowSpanEnd: spanObj.rowSpanEnd - 1,
          columnSpanEnd: spanObj.columnSpanEnd,
        );
        updateSpanCell = true;
      } else if (rowIndex <= spanObj.rowSpanEnd) {
        // Shrink the span as the row is removed.
        if (spanObj.rowSpanStart == spanObj.rowSpanEnd) {
          // Both are same then it means that it will no longer span.
          _spanList[i] = null;
          updateSpanCell = true;
        } else {
          _spanList[i] = _Span(
            rowSpanStart: spanObj.rowSpanStart,
            columnSpanStart: spanObj.columnSpanStart,
            rowSpanEnd: spanObj.rowSpanEnd - 1,
            columnSpanEnd: spanObj.columnSpanEnd,
          );
          updateSpanCell = true;
        }
      }
    }

    if (updateSpanCell) {
      _cleanUpSpanMap();
    }

    if (_sheetData.containsKey(rowIndex)) {
      _sheetData.remove(rowIndex);
    }
    Map<int, Map<int, Data>> _map = <int, Map<int, Data>>{};
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (rowKey > rowIndex) {
        _map[rowKey - 1] = _sheetData[rowKey]!;
        _map[rowKey - 1]!.values.forEach((element) {
          element._rowIndex = rowKey - 1;
        });
      } else {
        _map[rowKey] = _sheetData[rowKey]!;
      }
    });

    _sheetData = _map;
    _countRowsAndColumns();
  }

  ///
  /// Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  void insertRow(int rowIndex) {
    _checkMaxRow(rowIndex);
    if (rowIndex < 0) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex <= spanObj.rowSpanStart) {
        // Shifting required
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart + 1,
          columnSpanStart: spanObj.columnSpanStart,
          rowSpanEnd: spanObj.rowSpanEnd + 1,
          columnSpanEnd: spanObj.columnSpanEnd,
        );
        updateSpanCell = true;
      } else if (rowIndex <= spanObj.rowSpanEnd) {
        // Expand the span as the row is inserted.
        _spanList[i] = _Span(
          rowSpanStart: spanObj.rowSpanStart,
          columnSpanStart: spanObj.columnSpanStart,
          rowSpanEnd: spanObj.rowSpanEnd + 1,
          columnSpanEnd: spanObj.columnSpanEnd,
        );
        updateSpanCell = true;
      }
    }

    if (updateSpanCell) {
      _cleanUpSpanMap();
    }

    Map<int, Map<int, Data>> _map = <int, Map<int, Data>>{};
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (rowKey >= rowIndex) {
        _map[rowKey + 1] = _sheetData[rowKey]!;
        _map[rowKey + 1]!.values.forEach((element) {
          element._rowIndex = rowKey + 1;
        });
      } else {
        _map[rowKey] = _sheetData[rowKey]!;
      }
    });

    _sheetData = _map;
    _countRowsAndColumns();
  }

  ///
  /// Appends [row] iterables just post the last filled `rowIndex`.
  ///
  void appendRow(List<CellValue?> row) {
    int targetRow = maxRows;
    insertRowIterables(row, targetRow);
  }

  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  /// Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  /// [startingColumn] tells from where we should start putting the [row] iterables
  ///
  /// [overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  /// [overwriteMergedCells] when set to [false] puts the cell value to next unique cell available by putting the value in merged cells only once and jumps to next unique cell.
  ///
  void insertRowIterables(List<CellValue?> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    _checkMaxRow(rowIndex);
    if (rowIndex < 0) {
      return;
    }

    if (startingColumn < 0) {
      startingColumn = 0;
    }

    int i = startingColumn;

    /// Getting the Spanned Objects List
    List<_Span> spanObjectList = _getSpannedObjects(rowIndex, startingColumn);

    row.forEach((CellValue? cellValue) {
      if (!overwriteMergedCells) {
        while (!_isInsideSpanObject(spanObjectList, i, rowIndex)) {
          i++;
        }
      }

      _checkMaxColumn(i);

      _putData(rowIndex, i, cellValue);

      i++;
    });
    _countRowsAndColumns();
  }

  ///
  /// Returns the `count` of replaced `source` with `target`
  ///
  /// `source` is Pattern which allows you to pass your custom `RegExp` or a `String` providing more control over it.
  ///
  /// optional argument `first` is used to replace the number of first earlier occurrences
  ///
  /// If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  ///       excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///       or
  ///
  ///       var mySheet = excel['mySheetName'];
  ///       mySheet.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  /// In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  /// Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(Pattern source, dynamic target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0;

    if (startingRow == -1) {
      startingRow = 0;
    }
    if (endingRow == -1) {
      endingRow = maxRows;
    }
    if (startingColumn == -1) {
      startingColumn = 0;
    }
    if (endingColumn == -1) {
      endingColumn = maxColumns;
    }

    if (startingRow > endingRow) {
      int temp = startingRow;
      startingRow = endingRow;
      endingRow = temp;
    }
    if (startingColumn > endingColumn) {
      int temp = startingColumn;
      startingColumn = endingColumn;
      endingColumn = temp;
    }

    for (int i = startingRow; i <= endingRow; i++) {
      if (_sheetData[i] != null) {
        for (int j = startingColumn; j <= endingColumn; j++) {
          if (_sheetData[i]![j] != null && _sheetData[i]![j]!.value != null) {
            String value = _sheetData[i]![j]!.value.toString();
            if (value.contains(source)) {
              if (first != -1 && replaceCount >= first) {
                return replaceCount;
              }
              _sheetData[i]![j]!.value =
                  TextCellValue(value.replaceAll(source, target.toString()));
              replaceCount++;
            }
          }
        }
      }
    }
    return replaceCount;
  }

  ///
  /// Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  /// ----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  /// Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(CellIndex cellIndex, CellValue? value,
      {CellStyle? cellStyle}) {
    int columnIndex = cellIndex.columnIndex;
    int rowIndex = cellIndex.rowIndex;
    _checkMaxColumn(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    if (_spanList.isNotEmpty) {
      (newRowIndex, newColumnIndex) = _isInsideSpanning(rowIndex, columnIndex);
    }

    _putData(newRowIndex, newColumnIndex, value);

    if (cellStyle != null) {
      _sheetData[newRowIndex]![newColumnIndex]!.cellStyle = cellStyle;
    }
  }

  ///
  /// returns `true` if the contents are successfully `cleared` else `false`.
  ///
  /// If the row is having any spanned-cells then it will not be cleared and hence returns `false`.
  ///
  bool clearRow(int rowIndex) {
    if (rowIndex < 0) {
      return false;
    }

    bool isNotInside = true;

    if (_sheetData[rowIndex] != null && _sheetData[rowIndex]!.isNotEmpty) {
      for (int i = 0; i < _spanList.length; i++) {
        _Span? spanObj = _spanList[i];
        if (spanObj == null) {
          continue;
        }
        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd) {
          isNotInside = false;
          break;
        }
      }

      if (isNotInside) {
        _sheetData[rowIndex]!.keys.toList().forEach((key) {
          _sheetData[rowIndex]![key] = Data.newData(this, rowIndex, key);
        });
      }
    }
    return isNotInside;
  }
}
