part of '../../excel_community.dart';

/// Excel's default hyperlink font color.
const _hyperlinkColorHex = '#0563C1';

/// A hyperlink on a cell or range (`<hyperlink>` in the worksheet).
///
/// External targets ([url]: web pages, `mailto:` addresses, files) are
/// stored as external relationships of the worksheet; internal targets
/// ([location]: a cell, range or defined name) live on the element itself.
///
/// ```dart
/// sheet.setHyperlink(CellIndex.indexByString('A1'),
///     Hyperlink.url('https://pub.dev', tooltip: 'Open pub.dev'));
/// sheet.setHyperlink(CellIndex.indexByString('A2'),
///     Hyperlink.cell('Summary', 'B4'), text: 'Go to summary');
/// ```
class Hyperlink extends Equatable {
  /// External target: `https://...`, `mailto:...`, or a file path.
  final String? url;

  /// Target inside the workbook (`'My Sheet'!A1`, a range or a defined
  /// name). Combined with [url] it is the anchor within that document.
  final String? location;

  /// Text shown when hovering over the link.
  final String? tooltip;

  /// Optional display string stored with the link. Spreadsheet apps show
  /// the cell value; this attribute is informational.
  final String? display;

  const Hyperlink({this.url, this.location, this.tooltip, this.display})
      : assert(url != null || location != null,
            'A hyperlink needs a url or a location');

  /// Link to a web page, file or any other external address.
  factory Hyperlink.url(String url, {String? tooltip, String? display}) =>
      Hyperlink(url: url, tooltip: tooltip, display: display);

  /// `mailto:` link, with an optional pre-filled [subject].
  factory Hyperlink.email(String address,
      {String? subject, String? tooltip, String? display}) {
    final query =
        subject == null ? '' : '?subject=${Uri.encodeComponent(subject)}';
    return Hyperlink(
        url: 'mailto:$address$query', tooltip: tooltip, display: display);
  }

  /// Link to a cell or range of a worksheet in the same workbook. The sheet
  /// name is quoted when needed (`'Q1 Sales'!B4`).
  factory Hyperlink.cell(String sheetName, String cellReference,
          {String? tooltip, String? display}) =>
      Hyperlink(
        location: '${_quoteSheetName(sheetName)}!$cellReference',
        tooltip: tooltip,
        display: display,
      );

  /// Link to a defined name or a full reference such as `Sheet2!A1:C3`.
  factory Hyperlink.location(String location,
          {String? tooltip, String? display}) =>
      Hyperlink(location: location, tooltip: tooltip, display: display);

  /// Whether the target is outside the workbook.
  bool get isExternal => url != null;

  /// Text used for a cell that has no value yet.
  String get defaultText {
    if (display != null) return display!;
    final target = url ?? location!;
    return target.startsWith('mailto:')
        ? target.substring(7).split('?').first
        : target;
  }

  Hyperlink copyWith({
    String? url,
    String? location,
    String? tooltip,
    String? display,
  }) {
    return Hyperlink(
      url: url ?? this.url,
      location: location ?? this.location,
      tooltip: tooltip ?? this.tooltip,
      display: display ?? this.display,
    );
  }

  /// `<hyperlink>` element for [ref]; [relationshipId] links the external
  /// target registered in the worksheet relationships.
  String _toXmlString(String ref, String? relationshipId) {
    final sb = StringBuffer('<hyperlink ref="$ref"');
    if (relationshipId != null) sb.write(' r:id="$relationshipId"');
    if (location != null) sb.write(' location="${_escapeXml(location!)}"');
    if (tooltip != null) sb.write(' tooltip="${_escapeXml(tooltip!)}"');
    if (display != null) sb.write(' display="${_escapeXml(display!)}"');
    sb.write('/>');
    return sb.toString();
  }

  @override
  List<Object?> get props => [url, location, tooltip, display];

  @override
  String toString() =>
      'Hyperlink(${url ?? ''}${location != null ? '#$location' : ''})';
}

/// Quotes a sheet name for use in a reference when Excel would.
String _quoteSheetName(String name) {
  final plain = RegExp(r'^[A-Za-z_][A-Za-z0-9_.]*$').hasMatch(name) &&
      !RegExp(r'^[A-Za-z]{1,3}\d+$').hasMatch(name);
  return plain ? name : "'${name.replaceAll("'", "''")}'";
}

/// Hyperlinks of a worksheet.
extension SheetHyperlinks on Sheet {
  /// All hyperlinks keyed by cell or range reference (`'A1'`, `'B2:C3'`).
  Map<String, Hyperlink> get hyperlinks => Map.unmodifiable(_hyperlinks);

  /// Whether this worksheet has any hyperlink.
  bool get hasHyperlinks => _hyperlinks.isNotEmpty;

  /// Adds (or replaces) the hyperlink of [cellIndex].
  ///
  /// [text] sets the cell value; when the cell is empty and no [text] is
  /// given, [Hyperlink.defaultText] is written so the link is visible. With
  /// [styled] the cell gets Excel's hyperlink look (blue, underlined).
  void setHyperlink(CellIndex cellIndex, Hyperlink link,
      {String? text, bool styled = true}) {
    final cell = this.cell(cellIndex);
    if (text != null) {
      cell.value = TextCellValue(text);
    } else if (cell.value == null) {
      cell.value = TextCellValue(link.defaultText);
    }
    if (styled) _applyHyperlinkStyle(cell);
    _removeHyperlinksAt(cellIndex);
    _hyperlinks[cellIndex.cellId] = link;
  }

  /// Adds a hyperlink covering the whole [range] (e.g. `'A1:C1'`).
  void setHyperlinkRange(String range, Hyperlink link, {bool styled = true}) {
    final rect = _CellRect.parse(range);
    _hyperlinks.removeWhere((ref, _) => _CellRect.parse(ref).intersects(rect));
    if (styled) {
      for (var r = rect.top; r <= rect.bottom; r++) {
        for (var c = rect.left; c <= rect.right; c++) {
          _applyHyperlinkStyle(
              cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)));
        }
      }
    }
    _hyperlinks[rect.ref] = link;
  }

  /// The hyperlink on [cellIndex], directly or through a range.
  Hyperlink? getHyperlink(CellIndex cellIndex) {
    for (final entry in _hyperlinks.entries) {
      if (_CellRect.parse(entry.key).contains(cellIndex)) return entry.value;
    }
    return null;
  }

  /// Removes the hyperlink on [cellIndex] (including a range that covers
  /// it). The cell value and style are kept.
  void removeHyperlink(CellIndex cellIndex) => _removeHyperlinksAt(cellIndex);

  /// Removes every hyperlink of this worksheet.
  void clearHyperlinks() => _hyperlinks.clear();

  void _removeHyperlinksAt(CellIndex cellIndex) {
    _hyperlinks.removeWhere((ref, _) => _CellRect.parse(ref).contains(cellIndex));
  }

  void _applyHyperlinkStyle(Data cell) {
    final color = ExcelColor.fromHexString(_hyperlinkColorHex);
    final style = cell.cellStyle;
    cell.cellStyle = style == null
        ? CellStyle(fontColorHex: color, underline: Underline.Single)
        : style.copyWith(fontColorHexVal: color, underlineVal: Underline.Single);
  }

  /// Moves hyperlinks after rows/columns are inserted or removed. Links on a
  /// removed row/column are dropped.
  void _shiftHyperlinks({bool rows = true, required int index, required int delta}) {
    if (_hyperlinks.isEmpty) return;
    final shifted = <String, Hyperlink>{};
    _hyperlinks.forEach((ref, link) {
      final rect = _CellRect.parse(ref)
          .shifted(rows: rows, index: index, delta: delta);
      if (rect != null) shifted[rect.ref] = link;
    });
    _hyperlinks
      ..clear()
      ..addAll(shifted);
  }
}

/// Hyperlink access from a cell.
extension DataHyperlink on Data {
  /// The hyperlink on this cell, directly or through a range.
  Hyperlink? get hyperlink => _sheet.getHyperlink(cellIndex);

  /// Sets (or with `null` removes) the hyperlink of this cell, applying
  /// Excel's hyperlink style. See [SheetHyperlinks.setHyperlink].
  set hyperlink(Hyperlink? link) => link == null
      ? _sheet.removeHyperlink(cellIndex)
      : _sheet.setHyperlink(cellIndex, link);
}
