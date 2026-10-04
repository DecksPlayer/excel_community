part of '../../excel_community.dart';

/// Excel supports up to 7 nested outline levels.
const int _maxOutlineLevel = 7;

/// Outline (grouping) options of a worksheet (`<sheetPr><outlinePr>`).
class OutlineSettings extends Equatable {
  /// Summary rows are below their group (Excel's default); `false` puts
  /// them above, which moves the expand/collapse button to the top.
  final bool summaryBelow;

  /// Summary columns are to the right of their group (Excel's default).
  final bool summaryRight;

  /// Whether the outline bar with the +/- buttons is shown.
  final bool showOutlineSymbols;

  /// Whether automatic outline styles are applied.
  final bool applyStyles;

  const OutlineSettings({
    this.summaryBelow = true,
    this.summaryRight = true,
    this.showOutlineSymbols = true,
    this.applyStyles = false,
  });

  bool get isDefault =>
      summaryBelow && summaryRight && showOutlineSymbols && !applyStyles;

  OutlineSettings copyWith({
    bool? summaryBelow,
    bool? summaryRight,
    bool? showOutlineSymbols,
    bool? applyStyles,
  }) {
    return OutlineSettings(
      summaryBelow: summaryBelow ?? this.summaryBelow,
      summaryRight: summaryRight ?? this.summaryRight,
      showOutlineSymbols: showOutlineSymbols ?? this.showOutlineSymbols,
      applyStyles: applyStyles ?? this.applyStyles,
    );
  }

  /// `<outlinePr .../>`, or an empty string for the defaults.
  String toXmlString() {
    if (isDefault) return '';
    final sb = StringBuffer('<outlinePr');
    if (applyStyles) sb.write(' applyStyles="1"');
    if (!summaryBelow) sb.write(' summaryBelow="0"');
    if (!summaryRight) sb.write(' summaryRight="0"');
    if (!showOutlineSymbols) sb.write(' showOutlineSymbols="0"');
    sb.write('/>');
    return sb.toString();
  }

  @override
  List<Object?> get props =>
      [summaryBelow, summaryRight, showOutlineSymbols, applyStyles];
}

/// A run of consecutive rows (or columns) grouped at [level].
class OutlineGroup extends Equatable {
  /// First and last row/column index of the group (0-based, inclusive).
  final int start;
  final int end;

  /// Outline level (1-7); nested groups have higher levels.
  final int level;

  /// Whether the group is collapsed (its rows/columns are hidden).
  final bool collapsed;

  const OutlineGroup(this.start, this.end, this.level, {this.collapsed = false});

  @override
  List<Object?> get props => [start, end, level, collapsed];

  @override
  String toString() =>
      'OutlineGroup($start..$end, level $level${collapsed ? ', collapsed' : ''})';
}

/// Row and column grouping (Excel's Data > Group).
///
/// ```dart
/// sheet.groupRows(1, 4);                   // rows 2-5 under summary row 6
/// sheet.groupRows(2, 3);                   // nested group (level 2)
/// sheet.groupColumns(1, 3, collapsed: true);
/// sheet.expandRowGroup(1, 4);
/// ```
extension SheetGrouping on Sheet {
  /// Outline options; defaults to Excel's (summary rows below and summary
  /// columns to the right).
  OutlineSettings get outlineSettings =>
      _outlineSettings ?? const OutlineSettings();

  set outlineSettings(OutlineSettings settings) {
    _outlineSettings = settings.isDefault ? null : settings;
  }

  // ----- Rows ---------------------------------------------------------------

  /// Outline level of [rowIndex] (0 when not grouped).
  int getRowOutlineLevel(int rowIndex) => _rowOutlineLevels[rowIndex] ?? 0;

  /// Sets the outline level (0-7) of a single row.
  void setRowOutlineLevel(int rowIndex, int level) =>
      _setOutlineLevel(_rowOutlineLevels, rowIndex, level);

  /// Groups rows [start]..[end] (0-based, inclusive), nesting inside any
  /// existing group. With [collapsed] the group starts collapsed.
  ///
  /// Throws a [RangeError] when a row would exceed 7 levels.
  void groupRows(int start, int end, {bool collapsed = false}) {
    _group(_rowOutlineLevels, start, end);
    if (collapsed) collapseRowGroup(start, end);
  }

  /// Removes one outline level from rows [start]..[end], expanding the
  /// group first when it is collapsed.
  void ungroupRows(int start, int end) {
    if (_collapsedRows.contains(_rowSummary(start, end))) {
      expandRowGroup(start, end);
    }
    _ungroup(_rowOutlineLevels, start, end);
  }

  /// Hides rows [start]..[end] and marks their summary row as collapsed.
  void collapseRowGroup(int start, int end) => _collapse(
      _hiddenRows, _collapsedRows, start, end, _rowSummary(start, end));

  /// Shows rows [start]..[end] again; nested groups that are collapsed stay
  /// collapsed.
  void expandRowGroup(int start, int end) => _expand(_rowOutlineLevels,
      _hiddenRows, _collapsedRows, start, end, _rowSummary(start, end),
      rowGroups);

  /// All row groups, outer groups first.
  List<OutlineGroup> get rowGroups => _outlineGroups(
      _rowOutlineLevels, _collapsedRows, outlineSettings.summaryBelow);

  // ----- Columns --------------------------------------------------------------

  int getColumnOutlineLevel(int columnIndex) =>
      _columnOutlineLevels[columnIndex] ?? 0;

  void setColumnOutlineLevel(int columnIndex, int level) =>
      _setOutlineLevel(_columnOutlineLevels, columnIndex, level);

  /// Groups columns [start]..[end] (0-based, inclusive).
  void groupColumns(int start, int end, {bool collapsed = false}) {
    _group(_columnOutlineLevels, start, end);
    if (collapsed) collapseColumnGroup(start, end);
  }

  void ungroupColumns(int start, int end) {
    if (_collapsedColumns.contains(_columnSummary(start, end))) {
      expandColumnGroup(start, end);
    }
    _ungroup(_columnOutlineLevels, start, end);
  }

  void collapseColumnGroup(int start, int end) => _collapse(_hiddenColumns,
      _collapsedColumns, start, end, _columnSummary(start, end));

  void expandColumnGroup(int start, int end) => _expand(_columnOutlineLevels,
      _hiddenColumns, _collapsedColumns, start, end,
      _columnSummary(start, end), columnGroups);

  List<OutlineGroup> get columnGroups => _outlineGroups(
      _columnOutlineLevels, _collapsedColumns, outlineSettings.summaryRight);

  /// Removes every row and column group (rows/columns hidden by collapsed
  /// groups are shown again).
  void clearGrouping() {
    for (final group in rowGroups.where((g) => g.collapsed)) {
      expandRowGroup(group.start, group.end);
    }
    for (final group in columnGroups.where((g) => g.collapsed)) {
      expandColumnGroup(group.start, group.end);
    }
    _rowOutlineLevels.clear();
    _columnOutlineLevels.clear();
    _collapsedRows.clear();
    _collapsedColumns.clear();
  }

  // ----- Helpers ----------------------------------------------------------------

  int _rowSummary(int start, int end) =>
      outlineSettings.summaryBelow ? end + 1 : start - 1;

  int _columnSummary(int start, int end) =>
      outlineSettings.summaryRight ? end + 1 : start - 1;

  static void _checkSpan(int start, int end) {
    if (start < 0 || end < start) {
      throw RangeError('Invalid group span $start..$end');
    }
  }

  void _setOutlineLevel(Map<int, int> levels, int index, int level) {
    RangeError.checkNotNegative(index, 'index');
    RangeError.checkValueInInterval(level, 0, _maxOutlineLevel, 'level');
    if (level == 0) {
      levels.remove(index);
    } else {
      levels[index] = level;
    }
  }

  void _group(Map<int, int> levels, int start, int end) {
    _checkSpan(start, end);
    for (var i = start; i <= end; i++) {
      if ((levels[i] ?? 0) >= _maxOutlineLevel) {
        throw RangeError('Excel supports up to $_maxOutlineLevel outline levels');
      }
    }
    for (var i = start; i <= end; i++) {
      levels[i] = (levels[i] ?? 0) + 1;
    }
  }

  void _ungroup(Map<int, int> levels, int start, int end) {
    _checkSpan(start, end);
    for (var i = start; i <= end; i++) {
      final level = (levels[i] ?? 0) - 1;
      if (level <= 0) {
        levels.remove(i);
      } else {
        levels[i] = level;
      }
    }
  }

  void _collapse(
      Set<int> hidden, Set<int> collapsed, int start, int end, int summary) {
    _checkSpan(start, end);
    for (var i = start; i <= end; i++) {
      hidden.add(i);
    }
    if (summary >= 0) collapsed.add(summary);
  }

  void _expand(Map<int, int> levels, Set<int> hidden, Set<int> collapsed,
      int start, int end, int summary, List<OutlineGroup> groups) {
    _checkSpan(start, end);
    collapsed.remove(summary);
    var level = _maxOutlineLevel;
    for (var i = start; i <= end; i++) {
      level = min(level, levels[i] ?? 0);
    }
    // Rows of nested groups that are still collapsed stay hidden.
    final stillHidden = <int>{};
    for (final g in groups) {
      if (g.collapsed && g.level > level && g.start >= start && g.end <= end) {
        for (var i = g.start; i <= g.end; i++) {
          stillHidden.add(i);
        }
      }
    }
    for (var i = start; i <= end; i++) {
      if (!stillHidden.contains(i)) hidden.remove(i);
    }
  }
}

/// Groups formed by consecutive indexes with an outline level of at least
/// N, for every level N. The summary row/column is the one after the group
/// ([summaryAfter]) or before it.
List<OutlineGroup> _outlineGroups(
    Map<int, int> levels, Set<int> collapsed, bool summaryAfter) {
  if (levels.isEmpty) return const [];
  final indexes = levels.keys.toList()..sort();
  final maxLevel = levels.values.reduce(max);
  final groups = <OutlineGroup>[];
  for (var level = 1; level <= maxLevel; level++) {
    // Runs of consecutive indexes at this level or deeper.
    final runs = <(int, int)>[];
    for (final i in indexes) {
      if (levels[i]! < level) continue;
      if (runs.isNotEmpty && runs.last.$2 == i - 1) {
        runs.last = (runs.last.$1, i);
      } else {
        runs.add((i, i));
      }
    }
    for (final (start, end) in runs) {
      final summary = summaryAfter ? end + 1 : start - 1;
      groups.add(OutlineGroup(start, end, level,
          collapsed: collapsed.contains(summary)));
    }
  }
  groups.sort((a, b) =>
      a.start != b.start ? a.start.compareTo(b.start) : a.level.compareTo(b.level));
  return groups;
}

/// Moves index-keyed row/column properties after inserting (`delta: 1`) or
/// removing (`delta: -1`) the row/column at [index].
Map<int, V> _shiftIndexMap<V>(Map<int, V> map, int index, int delta) => {
      for (final entry in map.entries)
        if (!(delta < 0 && entry.key == index))
          (entry.key >= index ? entry.key + delta : entry.key): entry.value,
    };

Set<int> _shiftIndexSet(Set<int> set, int index, int delta) => {
      for (final i in set)
        if (!(delta < 0 && i == index)) (i >= index ? i + delta : i),
    };
