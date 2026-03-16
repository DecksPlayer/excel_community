part of excel_community;

extension SheetDimensions on Sheet {
  ///
  /// Set default column width
  ///
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultColumnWidth([double columnWidth = _excelDefaultColumnWidth]) {
    if (columnWidth < 0) return;
    _defaultColumnWidth = columnWidth;
  }

  ///
  /// Set default row height
  ///
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultRowHeight([double rowHeight = _excelDefaultRowHeight]) {
    if (rowHeight < 0) return;
    _defaultRowHeight = rowHeight;
  }

  ///
  /// Set Column AutoFit
  ///
  void setColumnAutoFit(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0) return;
    _columnAutoFit[columnIndex] = true;
  }

  ///
  /// Set Column Width
  ///
  void setColumnWidth(int columnIndex, double columnWidth) {
    _checkMaxColumn(columnIndex);
    if (columnWidth < 0) return;
    _columnWidths[columnIndex] = columnWidth;
  }

  ///
  /// Set Row Height
  ///
  void setRowHeight(int rowIndex, double rowHeight) {
    _checkMaxRow(rowIndex);
    if (rowHeight < 0) return;
    _rowHeights[rowIndex] = rowHeight;
  }

  ///
  /// returns auto fit state of column index
  ///
  bool getColumnAutoFit(int columnIndex) {
    if (_columnAutoFit.containsKey(columnIndex)) {
      return _columnAutoFit[columnIndex]!;
    }
    return false;
  }

  ///
  /// returns width of column index
  ///
  double getColumnWidth(int columnIndex) {
    if (_columnWidths.containsKey(columnIndex)) {
      return _columnWidths[columnIndex]!;
    }
    return _defaultColumnWidth!;
  }

  ///
  /// returns height of row index
  ///
  double getRowHeight(int rowIndex) {
    if (_rowHeights.containsKey(rowIndex)) {
      return _rowHeights[rowIndex]!;
    }
    return _defaultRowHeight!;
  }

  double? get defaultRowHeight => _defaultRowHeight;

  double? get defaultColumnWidth => _defaultColumnWidth;

  ///
  /// returns map of auto fit columns
  ///
  Map<int, bool> get getColumnAutoFits => _columnAutoFit;

  ///
  /// returns map of custom width columns
  ///
  Map<int, double> get getColumnWidths => _columnWidths;

  ///
  /// returns map of custom height rows
  ///
  Map<int, double> get getRowHeights => _rowHeights;
}
