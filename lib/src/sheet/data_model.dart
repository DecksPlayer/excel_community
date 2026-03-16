part of excel_community;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle? _cellStyle;
  CellValue? _value;
  Sheet _sheet;
  String _sheetName;
  int _rowIndex;
  int _columnIndex;

  ///
  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  ///
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.columnIndex,
          value: dataObject._value,
          cellStyleVal: dataObject._cellStyle,
        );

  ///
  ///Initializes the new `Data Object`
  ///
  Data._(
    Sheet sheet,
    int row,
    int column, {
    CellValue? value,
    NumFormat? numberFormat,
    CellStyle? cellStyleVal,
    bool isFormulaVal = false,
  })  : _sheet = sheet,
        _value = value,
        _cellStyle = cellStyleVal,
        _sheetName = sheet.sheetName,
        _rowIndex = row,
        _columnIndex = column;

  /// returns the newData object when called from Sheet Class
  static Data newData(Sheet sheet, int row, int column) {
    return Data._(sheet, row, column);
  }

  /// returns the row Index
  int get rowIndex {
    return _rowIndex;
  }

  /// returns the column Index
  int get columnIndex {
    return _columnIndex;
  }

  /// returns the sheet-name
  String get sheetName {
    return _sheetName;
  }

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
        columnIndex: _columnIndex, rowIndex: _rowIndex);
  }

  /// Helps to set the formula
  ///```
  ///var sheet = excel['Sheet1'];
  ///var cell = sheet.cell(CellIndex.indexByString("E5"));
  ///cell.setFormula('=SUM(1,2)');
  ///```
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, FormulaCellValue(formula));
  }

  set value(CellValue? val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  CellValue? get value => _value;

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? _) {
    _sheet._excel._styleChanges = true;
    _cellStyle = _;
  }

  @override
  List<Object?> get props => [
        _value,
        _columnIndex,
        _rowIndex,
        _cellStyle,
        _sheetName,
      ];
}
