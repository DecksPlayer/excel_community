part of '../../excel_community.dart';

/// Rectangle of cells (0-based, inclusive) used for range references such
/// as `A1` or `B2:D10` (hyperlinks, data validations).
class _CellRect {
  final int top;
  final int left;
  final int bottom;
  final int right;

  const _CellRect(this.top, this.left, this.bottom, this.right);

  /// Parses `A1` or `A1:C3` (`$` markers allowed, corners in any order).
  factory _CellRect.parse(String ref) {
    final parts = ref.trim().toUpperCase().replaceAll(r'$', '').split(':');
    final a = CellIndex.indexByString(parts.first);
    final b = parts.length > 1 ? CellIndex.indexByString(parts[1]) : a;
    return _CellRect(
      min(a.rowIndex, b.rowIndex),
      min(a.columnIndex, b.columnIndex),
      max(a.rowIndex, b.rowIndex),
      max(a.columnIndex, b.columnIndex),
    );
  }

  /// Parses a space-separated list of references (`sqref`).
  static List<_CellRect> parseList(String sqref) => sqref
      .split(RegExp(r'\s+'))
      .where((part) => part.isNotEmpty)
      .map(_CellRect.parse)
      .toList();

  CellIndex get start =>
      CellIndex.indexByColumnRow(columnIndex: left, rowIndex: top);
  CellIndex get end =>
      CellIndex.indexByColumnRow(columnIndex: right, rowIndex: bottom);

  /// `A1` for a single cell, `A1:C3` otherwise.
  String get ref => (top == bottom && left == right)
      ? start.cellId
      : '${start.cellId}:${end.cellId}';

  bool contains(CellIndex cell) =>
      cell.rowIndex >= top &&
      cell.rowIndex <= bottom &&
      cell.columnIndex >= left &&
      cell.columnIndex <= right;

  bool intersects(_CellRect other) =>
      other.left <= right &&
      other.right >= left &&
      other.top <= bottom &&
      other.bottom >= top;

  /// This rectangle minus [other], as up to four rectangles.
  List<_CellRect> subtract(_CellRect other) {
    if (!intersects(other)) return [this];
    final pieces = <_CellRect>[];
    if (other.top > top) pieces.add(_CellRect(top, left, other.top - 1, right));
    if (other.bottom < bottom) {
      pieces.add(_CellRect(other.bottom + 1, left, bottom, right));
    }
    final midTop = max(top, other.top);
    final midBottom = min(bottom, other.bottom);
    if (other.left > left) {
      pieces.add(_CellRect(midTop, left, midBottom, other.left - 1));
    }
    if (other.right < right) {
      pieces.add(_CellRect(midTop, other.right + 1, midBottom, right));
    }
    return pieces;
  }

  /// The rectangle after inserting (`delta: 1`) or removing (`delta: -1`)
  /// the row or column at [index]; `null` when a removal deletes it.
  _CellRect? shifted({required bool rows, required int index, required int delta}) {
    var lo = rows ? top : left;
    var hi = rows ? bottom : right;
    if (delta < 0) {
      if (lo == index && hi == index) return null;
      if (lo > index) lo += delta;
      if (hi >= index) hi += delta;
    } else {
      if (lo >= index) lo += delta;
      if (hi >= index) hi += delta;
    }
    return rows
        ? _CellRect(lo, left, hi, right)
        : _CellRect(top, lo, bottom, hi);
  }

  @override
  bool operator ==(Object other) =>
      other is _CellRect &&
      other.top == top &&
      other.left == left &&
      other.bottom == bottom &&
      other.right == right;

  @override
  int get hashCode => Object.hash(top, left, bottom, right);
}
