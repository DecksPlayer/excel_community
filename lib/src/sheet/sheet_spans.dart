part of excel_community;

extension SheetSpans on Sheet {
  ///
  /// returns List of Spanned Cells as
  ///
  ///     ["A1:A2", "A4:G6", "Y4:Y6", ....]
  ///
  ///return type if String based cell-id
  ///
  List<String> get spannedItems {
    _spannedItems = FastList<String>();

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      String rC = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
          spanObj.columnSpanEnd, spanObj.rowSpanEnd);
      if (!_spannedItems.contains(rC)) {
        _spannedItems.add(rC);
      }
    }

    return _spannedItems.keys;
  }

  ///
  /// Merges the cells starting from `start` to `end`.
  ///
  /// If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  void merge(CellIndex start, CellIndex end, {CellValue? customValue}) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    _checkMaxColumn(startColumn);
    _checkMaxColumn(endColumn);
    _checkMaxRow(startRow);
    _checkMaxRow(endRow);

    if ((startColumn == endColumn && startRow == endRow) ||
        (startColumn < 0 || startRow < 0 || endColumn < 0 || endRow < 0) ||
        (_spannedItems.contains(
            getSpanCellId(startColumn, startRow, endColumn, endRow)))) {
      return;
    }

    List<int> gotPosition = _getSpanPosition(start, end);

    _excel._mergeChanges = true;

    startColumn = gotPosition[0];
    startRow = gotPosition[1];
    endColumn = gotPosition[2];
    endRow = gotPosition[3];

    // Update maxColumns maxRows
    _maxColumns = _maxColumns > endColumn ? _maxColumns : endColumn + 1;
    _maxRows = _maxRows > endRow ? _maxRows : endRow + 1;

    bool getValue = true;

    Data value = Data.newData(this, startRow, startColumn);
    if (customValue != null) {
      value._value = customValue;
      getValue = false;
    }

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if (_sheetData[j] != null) {
          if (getValue && _sheetData[j]![k]?.value != null) {
            value = _sheetData[j]![k]!;
            getValue = false;
          }
          _sheetData[j]!.remove(k);
        }
      }
    }

    if (_sheetData[startRow] != null) {
      _sheetData[startRow]![startColumn] = value;
    } else {
      _sheetData[startRow] = {startColumn: value};
    }

    String sp = getSpanCellId(startColumn, startRow, endColumn, endRow);

    if (!_spannedItems.contains(sp)) {
      _spannedItems.add(sp);
    }

    _Span s = _Span(
      rowSpanStart: startRow,
      columnSpanStart: startColumn,
      rowSpanEnd: endRow,
      columnSpanEnd: endColumn,
    );

    _spanList.add(s);
    _excel._mergeChangeLookup = sheetName;
  }

  ///
  /// unMerge the merged cells.
  ///
  ///        var sheet = 'DesiredSheet';
  ///        List<String> spannedCells = excel.getMergedCells(sheet);
  ///        var cellToUnMerge = "A1:A2";
  ///        excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String unmergeCells) {
    if (_spannedItems.isNotEmpty &&
        _spanList.isNotEmpty &&
        _spannedItems.contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        CellIndex start = CellIndex.indexByString(lis[0]),
            end = CellIndex.indexByString(lis[1]);
        for (int i = 0; i < _spanList.length; i++) {
          _Span? spanObject = _spanList[i];
          if (spanObject == null) {
            continue;
          }

          if (spanObject.columnSpanStart == start.columnIndex &&
              spanObject.rowSpanStart == start.rowIndex &&
              spanObject.columnSpanEnd == end.columnIndex &&
              spanObject.rowSpanEnd == end.rowIndex) {
            _spanList[i] = null;
            remove = true;
          }
        }
        if (remove) {
          _cleanUpSpanMap();
        }
      }
      _spannedItems.remove(unmergeCells);
      _excel._mergeChangeLookup = sheetName;
    }
  }

  ///
  /// Sets the cellStyle of the merged cells.
  ///
  /// It will get the merged cells only by giving the starting position of merged cells.
  ///
  void setMergedCellStyle(CellIndex start, CellStyle mergedCellStyle) {
    List<List<CellIndex>> _mergedCells = spannedItems
        .map(
          (e) => e.split(":").map((e) => CellIndex.indexByString(e)).toList(),
        )
        .toList();

    List<CellIndex> _startIndices = _mergedCells.map((e) => e[0]).toList();
    List<CellIndex> _endIndices = _mergedCells.map((e) => e[1]).toList();

    if (_mergedCells.isEmpty ||
        start.columnIndex < 0 ||
        start.rowIndex < 0 ||
        !_startIndices.contains(start)) {
      return;
    }

    CellIndex end = _endIndices[_startIndices.indexOf(start)];

    bool hasBorder = mergedCellStyle.topBorder != Border() ||
        mergedCellStyle.bottomBorder != Border() ||
        mergedCellStyle.leftBorder != Border() ||
        mergedCellStyle.rightBorder != Border() ||
        mergedCellStyle.diagonalBorderUp ||
        mergedCellStyle.diagonalBorderDown;
    if (hasBorder) {
      for (var i = start.rowIndex; i <= end.rowIndex; i++) {
        for (var j = start.columnIndex; j <= end.columnIndex; j++) {
          CellStyle cellStyle = mergedCellStyle.copyWith(
            topBorderVal: Border(),
            bottomBorderVal: Border(),
            leftBorderVal: Border(),
            rightBorderVal: Border(),
            diagonalBorderUpVal: false,
            diagonalBorderDownVal: false,
          );

          if (i == start.rowIndex) {
            cellStyle = cellStyle.copyWith(
              topBorderVal: mergedCellStyle.topBorder,
            );
          }
          if (i == end.rowIndex) {
            cellStyle = cellStyle.copyWith(
              bottomBorderVal: mergedCellStyle.bottomBorder,
            );
          }
          if (j == start.columnIndex) {
            cellStyle = cellStyle.copyWith(
              leftBorderVal: mergedCellStyle.leftBorder,
            );
          }
          if (j == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              rightBorderVal: mergedCellStyle.rightBorder,
            );
          }

          if (i == j ||
              start.rowIndex == end.rowIndex ||
              start.columnIndex == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              diagonalBorderUpVal: mergedCellStyle.diagonalBorderUp,
              diagonalBorderDownVal: mergedCellStyle.diagonalBorderDown,
            );
          }

          if (i == start.rowIndex && j == start.columnIndex) {
            cell(start).cellStyle = cellStyle;
          } else {
            _putData(i, j, null);
            _sheetData[i]![j]!.cellStyle = cellStyle;
          }
        }
      }
    }
  }

  ///
  /// Helps to find the interaction between the pre-existing span position and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
  ///
  List<int> _getSpanPosition(CellIndex start, CellIndex end) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end.rowIndex;
      endRow = start.rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start.columnIndex;
      startColumn = end.columnIndex;
    }

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      final locationChange = _isLocationChangeRequired(
          startColumn, startRow, endColumn, endRow, spanObj);

      if (locationChange.$1) {
        startColumn = locationChange.$2.$1;
        startRow = locationChange.$2.$2;
        endColumn = locationChange.$2.$3;
        endRow = locationChange.$2.$4;
        String sp = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
            spanObj.columnSpanEnd, spanObj.rowSpanEnd);
        if (_spannedItems.contains(sp)) {
          _spannedItems.remove(sp);
        }
        remove = true;
        _spanList[i] = null;
      }
    }
    if (remove) {
      _cleanUpSpanMap();
    }

    return [startColumn, startRow, endColumn, endRow];
  }

  /// getting the List of _Span Objects which have the rowIndex containing and
  /// also lower the range by giving the starting columnIndex
  List<_Span> _getSpannedObjects(int rowIndex, int startingColumnIndex) {
    List<_Span> obtained = <_Span>[];

    if (_spanList.isNotEmpty) {
      obtained = <_Span>[];
      _spanList.forEach((spanObject) {
        if (spanObject != null &&
            spanObject.rowSpanStart <= rowIndex &&
            rowIndex <= spanObject.rowSpanEnd &&
            startingColumnIndex <= spanObject.columnSpanEnd) {
          obtained.add(spanObject);
        }
      });
    }
    return obtained;
  }

  ///
  /// Checking if the columnIndex and the rowIndex passed is inside the spanObjectList which is got from calling function.
  ///
  bool _isInsideSpanObject(
      List<_Span> spanObjectList, int columnIndex, int rowIndex) {
    for (int i = 0; i < spanObjectList.length; i++) {
      _Span spanObject = spanObjectList[i];

      if (spanObject.columnSpanStart <= columnIndex &&
          columnIndex <= spanObject.columnSpanEnd &&
          spanObject.rowSpanStart <= rowIndex &&
          rowIndex <= spanObject.rowSpanEnd) {
        if (columnIndex < spanObject.columnSpanEnd) {
          return false;
        } else if (columnIndex == spanObject.columnSpanEnd) {
          return true;
        }
      }
    }
    return true;
  }

  ///
  ///Cleans the `_SpanList` by removing the indexes where null value exists.
  ///
  void _cleanUpSpanMap() {
    if (_spanList.isNotEmpty) {
      _spanList.removeWhere((value) {
        return value == null;
      });
    }
  }
}
