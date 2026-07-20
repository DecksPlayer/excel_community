part of '../../excel_community.dart';

/// Enum representing standard OpenXML conditional formatting rule types.
enum ConditionalFormattingType {
  cellIs('cellIs'),
  expression('expression'),
  containsText('containsText'),
  notContains('notContainsText'),
  beginsWith('beginsWith'),
  endsWith('endsWith'),
  duplicateValues('duplicateValues'),
  uniqueValues('uniqueValues');

  final String value;
  const ConditionalFormattingType(this.value);

  static ConditionalFormattingType fromValue(String value) {
    return ConditionalFormattingType.values.firstWhere(
      (e) => e.value.toLowerCase() == value.toLowerCase(),
      orElse: () => ConditionalFormattingType.cellIs,
    );
  }
}

/// Enum representing comparison operators for [ConditionalFormattingType.cellIs] rules.
enum ConditionalFormattingOperator {
  equal('equal'),
  notEqual('notEqual'),
  greaterThan('greaterThan'),
  greaterThanOrEqual('greaterThanOrEqual'),
  lessThan('lessThan'),
  lessThanOrEqual('lessThanOrEqual'),
  between('between'),
  notBetween('notBetween'),
  containsText('containsText'),
  notContains('notContains'),
  beginsWith('beginsWith'),
  endsWith('endsWith');

  final String value;
  const ConditionalFormattingOperator(this.value);

  static ConditionalFormattingOperator? fromValue(String? value) {
    if (value == null) return null;
    return ConditionalFormattingOperator.values.firstWhereOrNull(
      (e) => e.value.toLowerCase() == value.toLowerCase(),
    );
  }
}

/// Differential style representation used in OpenXML `<dxf>` elements
/// for conditional formatting styling.
class DifferentialStyle extends Equatable {
  final ExcelColor? backgroundColor;
  final ExcelColor? fontColor;
  final bool? bold;
  final bool? italic;
  final bool? strikethrough;
  final Underline? underline;

  const DifferentialStyle({
    this.backgroundColor,
    this.fontColor,
    this.bold,
    this.italic,
    this.strikethrough,
    this.underline,
  });

  bool get isEmpty =>
      backgroundColor == null &&
      fontColor == null &&
      bold == null &&
      italic == null &&
      strikethrough == null &&
      (underline == null || underline == Underline.None);

  @override
  List<Object?> get props => [
        backgroundColor?.colorHex,
        fontColor?.colorHex,
        bold,
        italic,
        strikethrough,
        underline,
      ];
}

/// Represents a single rule (`<cfRule>`) inside a conditional formatting group.
class ConditionalFormattingRule {
  final ConditionalFormattingType type;
  final ConditionalFormattingOperator? operator;
  final List<String> formulae;
  final DifferentialStyle style;
  final int priority;
  final String? text;

  ConditionalFormattingRule({
    required this.type,
    this.operator,
    List<String>? formulae,
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    this.priority = 1,
    this.text,
  })  : formulae = formulae ?? const [],
        style = style ??
            DifferentialStyle(
              backgroundColor: backgroundColor,
              fontColor: fontColor,
              bold: bold,
              italic: italic,
              strikethrough: strikethrough,
              underline: underline,
            );

  /// Factory for a cell comparison rule (e.g. greater than, equal, between).
  factory ConditionalFormattingRule.cellIs({
    required ConditionalFormattingOperator operator,
    required List<String> formulae,
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    int priority = 1,
  }) {
    return ConditionalFormattingRule(
      type: ConditionalFormattingType.cellIs,
      operator: operator,
      formulae: formulae,
      style: style,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      priority: priority,
    );
  }

  /// Factory for a rule matching specific text in cells.
  factory ConditionalFormattingRule.containsText({
    required String text,
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    int priority = 1,
  }) {
    return ConditionalFormattingRule(
      type: ConditionalFormattingType.containsText,
      operator: ConditionalFormattingOperator.containsText,
      text: text,
      style: style,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      priority: priority,
    );
  }

  /// Factory for a rule with a custom formula/expression.
  factory ConditionalFormattingRule.expression({
    required String formula,
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    int priority = 1,
  }) {
    return ConditionalFormattingRule(
      type: ConditionalFormattingType.expression,
      formulae: [formula],
      style: style,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      priority: priority,
    );
  }

  /// Factory for highlighting duplicate values in a range.
  factory ConditionalFormattingRule.duplicateValues({
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    int priority = 1,
  }) {
    return ConditionalFormattingRule(
      type: ConditionalFormattingType.duplicateValues,
      style: style,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      priority: priority,
    );
  }

  /// Factory for highlighting unique values in a range.
  factory ConditionalFormattingRule.uniqueValues({
    DifferentialStyle? style,
    ExcelColor? backgroundColor,
    ExcelColor? fontColor,
    bool? bold,
    bool? italic,
    bool? strikethrough,
    Underline? underline,
    int priority = 1,
  }) {
    return ConditionalFormattingRule(
      type: ConditionalFormattingType.uniqueValues,
      style: style,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      priority: priority,
    );
  }
}

/// Represents a collection of rules associated with a specific cell range (`sqref`).
class ConditionalFormattingGroup {
  final String sqref;
  final List<ConditionalFormattingRule> rules;

  ConditionalFormattingGroup({
    required this.sqref,
    required List<ConditionalFormattingRule> rules,
  }) : rules = List.unmodifiable(rules);
}
