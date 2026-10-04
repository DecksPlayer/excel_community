part of '../../excel_community.dart';

/// Kind of value a [DataValidation] allows (`<dataValidation type="...">`).
enum DataValidationType {
  /// Any value (used to show an input message only).
  any('none'),
  wholeNumber('whole'),
  decimal('decimal'),

  /// One value of a list: an in-cell dropdown.
  list('list'),
  date('date'),
  time('time'),
  textLength('textLength'),

  /// A formula that must evaluate to TRUE.
  custom('custom');

  final String xmlValue;
  const DataValidationType(this.xmlValue);

  static DataValidationType fromXmlValue(String? value) =>
      values.firstWhere((t) => t.xmlValue == value, orElse: () => any);
}

/// Comparison used by number, date, time and text length validations.
enum DataValidationOperator {
  between('between'),
  notBetween('notBetween'),
  equal('equal'),
  notEqual('notEqual'),
  lessThan('lessThan'),
  lessThanOrEqual('lessThanOrEqual'),
  greaterThan('greaterThan'),
  greaterThanOrEqual('greaterThanOrEqual');

  final String xmlValue;
  const DataValidationOperator(this.xmlValue);

  /// Whether the operator needs two values.
  bool get isRange => this == between || this == notBetween;

  static DataValidationOperator fromXmlValue(String? value) =>
      values.firstWhere((o) => o.xmlValue == value, orElse: () => between);
}

/// How Excel reacts to an invalid value.
enum DataValidationErrorStyle {
  /// Rejects the value.
  stop('stop'),

  /// Asks the user to confirm (Yes/No/Cancel).
  warning('warning'),

  /// Informs the user and accepts the value.
  information('information');

  final String xmlValue;
  const DataValidationErrorStyle(this.xmlValue);

  static DataValidationErrorStyle fromXmlValue(String? value) =>
      values.firstWhere((s) => s.xmlValue == value, orElse: () => stop);
}

/// A data validation rule (`<dataValidation>`): dropdown lists, number,
/// date, time and text length limits or custom formulas, with optional
/// input and error messages.
///
/// ```dart
/// sheet.addDataValidation('C2:C100',
///     DataValidation.list(['Open', 'In progress', 'Done'])
///         .withPrompt('Status', 'Pick a status from the list')
///         .withError('Invalid status', 'Choose one of the listed values'));
///
/// sheet.addDataValidation('D2:D100',
///     DataValidation.wholeNumber(DataValidationOperator.between, 1, 10));
/// ```
class DataValidation extends Equatable {
  final DataValidationType type;
  final DataValidationOperator operator;

  /// First value, list source or formula (without the leading `=`).
  final String? formula1;

  /// Second value for [DataValidationOperator.between] / `notBetween`.
  final String? formula2;

  /// Whether empty cells are valid.
  final bool allowBlank;

  /// Whether list validations show the in-cell dropdown arrow.
  final bool showDropdown;

  final bool showInputMessage;
  final bool showErrorMessage;
  final String? promptTitle;
  final String? prompt;
  final String? errorTitle;
  final String? error;
  final DataValidationErrorStyle errorStyle;

  const DataValidation({
    required this.type,
    this.operator = DataValidationOperator.between,
    this.formula1,
    this.formula2,
    this.allowBlank = true,
    this.showDropdown = true,
    this.showInputMessage = true,
    this.showErrorMessage = true,
    this.promptTitle,
    this.prompt,
    this.errorTitle,
    this.error,
    this.errorStyle = DataValidationErrorStyle.stop,
  });

  /// Dropdown with fixed [items]. Items cannot contain commas or double
  /// quotes, and the joined list is limited to 255 characters by Excel.
  factory DataValidation.list(List<String> items) {
    if (items.isEmpty) {
      throw ArgumentError.value(items, 'items', 'must not be empty');
    }
    for (final item in items) {
      if (item.contains(',') || item.contains('"')) {
        throw ArgumentError.value(item, 'items',
            'must not contain commas or quotes; use DataValidation.listFromRange');
      }
    }
    final joined = items.join(',');
    if (joined.length > 255) {
      throw ArgumentError.value(items, 'items',
          'exceed 255 characters; use DataValidation.listFromRange');
    }
    return DataValidation(type: DataValidationType.list, formula1: '"$joined"');
  }

  /// Dropdown with the values of a cell range, e.g. `'A1:A10'`, optionally
  /// on another [sheetName] (references are made absolute).
  factory DataValidation.listFromRange(String range, {String? sheetName}) {
    final rect = _CellRect.parse(range);
    String absolute(CellIndex c) =>
        '\$${getColumnAlphabet(c.columnIndex)}\$${c.rowIndex + 1}';
    final ref = rect.top == rect.bottom && rect.left == rect.right
        ? absolute(rect.start)
        : '${absolute(rect.start)}:${absolute(rect.end)}';
    return DataValidation(
      type: DataValidationType.list,
      formula1: sheetName == null ? ref : '${_quoteSheetName(sheetName)}!$ref',
    );
  }

  /// Whole number compared with [value] (and [value2] for between).
  factory DataValidation.wholeNumber(DataValidationOperator operator, int value,
          [int? value2]) =>
      _compare(DataValidationType.wholeNumber, operator, '$value',
          value2?.toString());

  /// Decimal number compared with [value] (and [value2] for between).
  factory DataValidation.decimal(DataValidationOperator operator, num value,
          [num? value2]) =>
      _compare(DataValidationType.decimal, operator, _number(value),
          value2 == null ? null : _number(value2));

  /// Date compared with [value] (and [value2] for between).
  factory DataValidation.date(DataValidationOperator operator, DateTime value,
          [DateTime? value2]) =>
      _compare(DataValidationType.date, operator, '${_dateSerial(value)}',
          value2 == null ? null : '${_dateSerial(value2)}');

  /// Time of day compared with [value] (and [value2] for between).
  factory DataValidation.time(DataValidationOperator operator, Duration value,
          [Duration? value2]) =>
      _compare(DataValidationType.time, operator, _number(_dayFraction(value)),
          value2 == null ? null : _number(_dayFraction(value2)));

  /// Text length compared with [length] (and [length2] for between).
  factory DataValidation.textLength(DataValidationOperator operator, int length,
          [int? length2]) =>
      _compare(DataValidationType.textLength, operator, '$length',
          length2?.toString());

  /// Valid when [formula] (without `=`) is TRUE for the cell, e.g.
  /// `'ISNUMBER(A2)'` for a rule applied to `A2:A100`.
  factory DataValidation.custom(String formula) => DataValidation(
      type: DataValidationType.custom,
      formula1: formula.startsWith('=') ? formula.substring(1) : formula);

  /// No restriction: only shows the input message.
  factory DataValidation.inputMessage(String title, String message) =>
      DataValidation(
          type: DataValidationType.any, promptTitle: title, prompt: message);

  static DataValidation _compare(DataValidationType type,
      DataValidationOperator operator, String value, String? value2) {
    if (operator.isRange && value2 == null) {
      throw ArgumentError('${operator.xmlValue} needs a second value');
    }
    return DataValidation(
      type: type,
      operator: operator,
      formula1: value,
      formula2: operator.isRange ? value2 : null,
    );
  }

  /// Copy with the input message shown when the cell is selected.
  DataValidation withPrompt(String title, String message) =>
      copyWith(promptTitle: title, prompt: message, showInputMessage: true);

  /// Copy with the message shown for an invalid value.
  DataValidation withError(String title, String message,
          {DataValidationErrorStyle style = DataValidationErrorStyle.stop}) =>
      copyWith(
          errorTitle: title,
          error: message,
          errorStyle: style,
          showErrorMessage: true);

  /// Items of a fixed-list validation (`null` for range or formula lists).
  List<String>? get listItems {
    final source = formula1;
    if (type != DataValidationType.list ||
        source == null ||
        !source.startsWith('"') ||
        !source.endsWith('"')) {
      return null;
    }
    return source.substring(1, source.length - 1).split(',');
  }

  /// Checks [value] against this rule like Excel would.
  ///
  /// Returns `null` when the rule depends on formulas, cell ranges or other
  /// cells and cannot be evaluated here.
  bool? accepts(CellValue? value) {
    final resolved = value is FormulaCellValue ? value.cachedValue : value;
    if (resolved == null ||
        (resolved is TextCellValue && resolved.toString().isEmpty)) {
      return allowBlank;
    }
    switch (type) {
      case DataValidationType.any:
        return true;
      case DataValidationType.custom:
        return null;
      case DataValidationType.list:
        final items = listItems;
        if (items == null) return null;
        final text = _valueText(resolved).toLowerCase();
        return items.any((item) => item.trim().toLowerCase() == text);
      case DataValidationType.textLength:
        return _check(_valueText(resolved).length);
      case DataValidationType.wholeNumber:
        final n = _numeric(resolved);
        if (n == null || n != n.roundToDouble()) return false;
        return _check(n);
      case DataValidationType.decimal:
        final n = _numeric(resolved);
        return n == null ? false : _check(n);
      case DataValidationType.date:
        final serial = switch (resolved) {
          DateCellValue v => _dateSerial(v.asDateTimeUtc()).toDouble(),
          DateTimeCellValue v => _dateSerial(v.asDateTimeUtc()) +
              _dayFraction(Duration(
                  hours: v.hour, minutes: v.minute, seconds: v.second)),
          _ => null,
        };
        return serial == null ? false : _check(serial);
      case DataValidationType.time:
        final fraction = switch (resolved) {
          TimeCellValue v => _dayFraction(v.asDuration()),
          DateTimeCellValue v => _dayFraction(
              Duration(hours: v.hour, minutes: v.minute, seconds: v.second)),
          _ => null,
        };
        return fraction == null ? false : _check(fraction);
    }
  }

  bool? _check(num value) {
    final a = double.tryParse(formula1 ?? '');
    if (a == null) return null;
    final b = double.tryParse(formula2 ?? '');
    if (operator.isRange && b == null) return null;
    return switch (operator) {
      DataValidationOperator.between => value >= a && value <= b!,
      DataValidationOperator.notBetween => value < a || value > b!,
      DataValidationOperator.equal => value == a,
      DataValidationOperator.notEqual => value != a,
      DataValidationOperator.lessThan => value < a,
      DataValidationOperator.lessThanOrEqual => value <= a,
      DataValidationOperator.greaterThan => value > a,
      DataValidationOperator.greaterThanOrEqual => value >= a,
    };
  }

  static String _valueText(CellValue value) => switch (value) {
        BoolCellValue v => v.value ? 'TRUE' : 'FALSE',
        _ => value.toString(),
      };

  static double? _numeric(CellValue value) => switch (value) {
        IntCellValue v => v.value.toDouble(),
        DoubleCellValue v => v.value,
        _ => null,
      };

  DataValidation copyWith({
    DataValidationType? type,
    DataValidationOperator? operator,
    String? formula1,
    String? formula2,
    bool? allowBlank,
    bool? showDropdown,
    bool? showInputMessage,
    bool? showErrorMessage,
    String? promptTitle,
    String? prompt,
    String? errorTitle,
    String? error,
    DataValidationErrorStyle? errorStyle,
  }) {
    return DataValidation(
      type: type ?? this.type,
      operator: operator ?? this.operator,
      formula1: formula1 ?? this.formula1,
      formula2: formula2 ?? this.formula2,
      allowBlank: allowBlank ?? this.allowBlank,
      showDropdown: showDropdown ?? this.showDropdown,
      showInputMessage: showInputMessage ?? this.showInputMessage,
      showErrorMessage: showErrorMessage ?? this.showErrorMessage,
      promptTitle: promptTitle ?? this.promptTitle,
      prompt: prompt ?? this.prompt,
      errorTitle: errorTitle ?? this.errorTitle,
      error: error ?? this.error,
      errorStyle: errorStyle ?? this.errorStyle,
    );
  }

  /// `<dataValidation>` element for the space-separated ranges [sqref].
  String _toXmlString(String sqref) {
    final sb = StringBuffer('<dataValidation');
    if (type != DataValidationType.any) sb.write(' type="${type.xmlValue}"');
    if (errorStyle != DataValidationErrorStyle.stop) {
      sb.write(' errorStyle="${errorStyle.xmlValue}"');
    }
    if (operator != DataValidationOperator.between) {
      sb.write(' operator="${operator.xmlValue}"');
    }
    if (allowBlank) sb.write(' allowBlank="1"');
    // Note the inverted meaning: showDropDown="1" hides the arrow.
    if (!showDropdown) sb.write(' showDropDown="1"');
    if (showInputMessage) sb.write(' showInputMessage="1"');
    if (showErrorMessage) sb.write(' showErrorMessage="1"');
    if (errorTitle != null) sb.write(' errorTitle="${_escapeXml(errorTitle!)}"');
    if (error != null) sb.write(' error="${_escapeXml(error!)}"');
    if (promptTitle != null) {
      sb.write(' promptTitle="${_escapeXml(promptTitle!)}"');
    }
    if (prompt != null) sb.write(' prompt="${_escapeXml(prompt!)}"');
    sb.write(' sqref="$sqref">');
    if (formula1 != null) sb.write('<formula1>${_escapeXml(formula1!)}</formula1>');
    if (formula2 != null) sb.write('<formula2>${_escapeXml(formula2!)}</formula2>');
    sb.write('</dataValidation>');
    return sb.toString();
  }

  @override
  List<Object?> get props => [
        type,
        operator,
        formula1,
        formula2,
        allowBlank,
        showDropdown,
        showInputMessage,
        showErrorMessage,
        promptTitle,
        prompt,
        errorTitle,
        error,
        errorStyle,
      ];
}

/// Excel serial day number of [date] (days since 1899-12-30).
int _dateSerial(DateTime date) => DateTime.utc(date.year, date.month, date.day)
    .difference(DateTime.utc(1899, 12, 30))
    .inDays;

double _dayFraction(Duration time) =>
    time.inMilliseconds / Duration.millisecondsPerDay;

String _number(num value) {
  final text = value.toString();
  return text.endsWith('.0') ? text.substring(0, text.length - 2) : text;
}

/// Data validations of a worksheet.
extension SheetDataValidations on Sheet {
  /// Validations keyed by their space-separated ranges (`'A2:A10 C2:C10'`).
  Map<String, DataValidation> get dataValidations =>
      Map.unmodifiable(_dataValidations);

  bool get hasDataValidations => _dataValidations.isNotEmpty;

  /// Applies [validation] to [range] (`'A2:A100'`, or several ranges
  /// separated by spaces). A cell has at most one validation, so the area is
  /// removed from any validation that already covered it.
  void addDataValidation(String range, DataValidation validation) {
    final rects = _CellRect.parseList(range);
    if (rects.isEmpty) throw ArgumentError.value(range, 'range');
    _subtractValidationArea(rects);
    _dataValidations[rects.map((r) => r.ref).join(' ')] = validation;
  }

  /// Applies [validation] to a single cell.
  void setDataValidation(CellIndex cellIndex, DataValidation validation) =>
      addDataValidation(cellIndex.cellId, validation);

  /// The validation that applies to [cellIndex], if any.
  DataValidation? getDataValidation(CellIndex cellIndex) {
    for (final entry in _dataValidations.entries) {
      if (_CellRect.parseList(entry.key).any((r) => r.contains(cellIndex))) {
        return entry.value;
      }
    }
    return null;
  }

  /// Removes validation from [range]; validations partly inside it keep the
  /// rest of their area.
  void removeDataValidation(String range) =>
      _subtractValidationArea(_CellRect.parseList(range));

  void clearDataValidations() => _dataValidations.clear();

  void _subtractValidationArea(List<_CellRect> area) {
    final updated = <String, DataValidation>{};
    _dataValidations.forEach((sqref, validation) {
      var rects = _CellRect.parseList(sqref);
      for (final cut in area) {
        rects = [for (final r in rects) ...r.subtract(cut)];
      }
      if (rects.isNotEmpty) updated[rects.map((r) => r.ref).join(' ')] = validation;
    });
    _dataValidations
      ..clear()
      ..addAll(updated);
  }

  void _shiftDataValidations({bool rows = true, required int index, required int delta}) {
    if (_dataValidations.isEmpty) return;
    final shifted = <String, DataValidation>{};
    _dataValidations.forEach((sqref, validation) {
      final rects = [
        for (final r in _CellRect.parseList(sqref))
          if (r.shifted(rows: rows, index: index, delta: delta) case final s?) s,
      ];
      if (rects.isNotEmpty) shifted[rects.map((r) => r.ref).join(' ')] = validation;
    });
    _dataValidations
      ..clear()
      ..addAll(shifted);
  }
}

/// Data validation access from a cell.
extension DataDataValidation on Data {
  /// The validation applying to this cell, if any.
  DataValidation? get dataValidation => _sheet.getDataValidation(cellIndex);

  /// Sets (or with `null` removes) the validation of this cell.
  set dataValidation(DataValidation? validation) => validation == null
      ? _sheet.removeDataValidation(cellIndex.cellId)
      : _sheet.setDataValidation(cellIndex, validation);
}
