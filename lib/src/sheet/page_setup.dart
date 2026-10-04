part of '../../excel_community.dart';

/// Page orientation used when printing a worksheet (`<pageSetup orientation="...">`).
enum PageOrientation {
  /// Uses the printer's default orientation (OpenXML value `default`).
  automatic('default'),
  portrait('portrait'),
  landscape('landscape');

  final String xmlValue;
  const PageOrientation(this.xmlValue);

  static PageOrientation? fromXmlValue(String? value) {
    for (final o in PageOrientation.values) {
      if (o.xmlValue == value) return o;
    }
    return null;
  }
}

/// Order in which pages are numbered and printed when the sheet spans
/// several pages both horizontally and vertically.
enum PageOrder {
  downThenOver('downThenOver'),
  overThenDown('overThenDown');

  final String xmlValue;
  const PageOrder(this.xmlValue);

  static PageOrder? fromXmlValue(String? value) {
    for (final o in PageOrder.values) {
      if (o.xmlValue == value) return o;
    }
    return null;
  }
}

/// How cell comments are printed.
enum PrintCellComments {
  none('none'),
  atEnd('atEnd'),
  asDisplayed('asDisplayed');

  final String xmlValue;
  const PrintCellComments(this.xmlValue);

  static PrintCellComments? fromXmlValue(String? value) {
    for (final c in PrintCellComments.values) {
      if (c.xmlValue == value) return c;
    }
    return null;
  }
}

/// How cell errors (`#DIV/0!`, `#N/A`, ...) are printed.
enum PrintErrors {
  displayed('displayed'),
  blank('blank'),
  dash('dash'),
  na('NA');

  final String xmlValue;
  const PrintErrors(this.xmlValue);

  static PrintErrors? fromXmlValue(String? value) {
    for (final e in PrintErrors.values) {
      if (e.xmlValue == value) return e;
    }
    return null;
  }
}

/// Paper size used when printing, identified by its SpreadsheetML code
/// (`<pageSetup paperSize="9">` is A4).
///
/// Common sizes are exposed as constants. Any other code from the
/// ECMA-376 table can be used through [PaperSize.fromCode], and unknown codes
/// read from a file are preserved as-is.
class PaperSize extends Equatable {
  /// SpreadsheetML paper size code.
  final int code;

  /// Human readable name, e.g. `'A4'`.
  final String name;

  /// Paper width in millimetres, when known.
  final double? widthMm;

  /// Paper height in millimetres, when known.
  final double? heightMm;

  const PaperSize._(this.code, this.name, [this.widthMm, this.heightMm]);

  static const letter = PaperSize._(1, 'Letter', 215.9, 279.4);
  static const tabloid = PaperSize._(3, 'Tabloid', 279.4, 431.8);
  static const ledger = PaperSize._(4, 'Ledger', 431.8, 279.4);
  static const legal = PaperSize._(5, 'Legal', 215.9, 355.6);
  static const statement = PaperSize._(6, 'Statement', 139.7, 215.9);
  static const executive = PaperSize._(7, 'Executive', 184.15, 266.7);
  static const a3 = PaperSize._(8, 'A3', 297, 420);
  static const a4 = PaperSize._(9, 'A4', 210, 297);
  static const a5 = PaperSize._(11, 'A5', 148, 210);
  static const b4 = PaperSize._(12, 'B4 (JIS)', 257, 364);
  static const b5 = PaperSize._(13, 'B5 (JIS)', 182, 257);
  static const folio = PaperSize._(14, 'Folio', 215.9, 330.2);
  static const quarto = PaperSize._(15, 'Quarto', 215, 275);
  static const envelope10 = PaperSize._(20, 'Envelope #10', 104.8, 241.3);
  static const envelopeDL = PaperSize._(27, 'Envelope DL', 110, 220);
  static const envelopeC5 = PaperSize._(28, 'Envelope C5', 162, 229);
  static const a2 = PaperSize._(66, 'A2', 420, 594);
  static const a6 = PaperSize._(70, 'A6', 105, 148);

  /// All paper sizes with a named constant.
  static const List<PaperSize> values = [
    letter,
    tabloid,
    ledger,
    legal,
    statement,
    executive,
    a3,
    a4,
    a5,
    b4,
    b5,
    folio,
    quarto,
    envelope10,
    envelopeDL,
    envelopeC5,
    a2,
    a6,
  ];

  /// Returns the known [PaperSize] for [code], or a custom one that keeps
  /// the raw code so it round-trips unchanged.
  factory PaperSize.fromCode(int code) {
    for (final p in values) {
      if (p.code == code) return p;
    }
    return PaperSize._(code, 'Paper size $code');
  }

  @override
  List<Object?> get props => [code];

  @override
  String toString() => 'PaperSize($name, code: $code)';
}

/// Represents the OpenXML `<pageSetup>` element of a worksheet, plus the
/// `fitToPage` flag that Excel stores in `<sheetPr><pageSetUpPr>`.
///
/// Every property is optional; only the ones that are set are written, so
/// spreadsheet applications fall back to their own defaults for the rest.
///
/// ```dart
/// sheet.pageSetup = PageSetup(
///   orientation: PageOrientation.landscape,
///   paperSize: PaperSize.a4,
/// );
///
/// // Fit all columns on one page; as many pages tall as needed.
/// sheet.fitToPages(width: 1, height: 0);
/// ```
class PageSetup extends Equatable {
  /// Page orientation.
  final PageOrientation? orientation;

  /// Paper size.
  final PaperSize? paperSize;

  /// Custom paper width with units (e.g. `'210mm'`, `'8.5in'`). Overrides
  /// [paperSize] in applications that support it.
  final String? paperWidth;

  /// Custom paper height with units (e.g. `'297mm'`, `'11in'`).
  final String? paperHeight;

  /// Print scaling percentage (10–400). Ignored when [fitToPage] is `true`.
  final int? scale;

  /// When `true`, the sheet is scaled to fit [fitToWidth] × [fitToHeight]
  /// pages instead of using [scale]. Written to `<sheetPr><pageSetUpPr>`.
  final bool? fitToPage;

  /// Number of pages wide the sheet is fitted to (`0` = automatic).
  final int? fitToWidth;

  /// Number of pages tall the sheet is fitted to (`0` = automatic).
  final int? fitToHeight;

  /// First printed page number. Only used when [useFirstPageNumber] is `true`.
  final int? firstPageNumber;

  /// Whether [firstPageNumber] is applied.
  final bool? useFirstPageNumber;

  /// Page numbering/printing order.
  final PageOrder? pageOrder;

  /// Print in black and white.
  final bool? blackAndWhite;

  /// Print in draft quality.
  final bool? draft;

  /// How cell comments are printed.
  final PrintCellComments? cellComments;

  /// How cell errors are printed.
  final PrintErrors? errors;

  /// Horizontal print resolution in DPI.
  final int? horizontalDpi;

  /// Vertical print resolution in DPI.
  final int? verticalDpi;

  /// Number of copies to print.
  final int? copies;

  /// Whether the printer's defaults are used for unset values.
  final bool? usePrinterDefaults;

  const PageSetup({
    this.orientation,
    this.paperSize,
    this.paperWidth,
    this.paperHeight,
    this.scale,
    this.fitToPage,
    this.fitToWidth,
    this.fitToHeight,
    this.firstPageNumber,
    this.useFirstPageNumber,
    this.pageOrder,
    this.blackAndWhite,
    this.draft,
    this.cellComments,
    this.errors,
    this.horizontalDpi,
    this.verticalDpi,
    this.copies,
    this.usePrinterDefaults,
  });

  /// Whether any `<pageSetup>` attribute is set (ignores [fitToPage], which
  /// lives in `<sheetPr>`).
  bool get hasPageSetupAttributes =>
      orientation != null ||
      paperSize != null ||
      paperWidth != null ||
      paperHeight != null ||
      scale != null ||
      fitToWidth != null ||
      fitToHeight != null ||
      firstPageNumber != null ||
      useFirstPageNumber != null ||
      pageOrder != null ||
      blackAndWhite != null ||
      draft != null ||
      cellComments != null ||
      errors != null ||
      horizontalDpi != null ||
      verticalDpi != null ||
      copies != null ||
      usePrinterDefaults != null;

  PageSetup copyWith({
    PageOrientation? orientation,
    PaperSize? paperSize,
    String? paperWidth,
    String? paperHeight,
    int? scale,
    bool? fitToPage,
    int? fitToWidth,
    int? fitToHeight,
    int? firstPageNumber,
    bool? useFirstPageNumber,
    PageOrder? pageOrder,
    bool? blackAndWhite,
    bool? draft,
    PrintCellComments? cellComments,
    PrintErrors? errors,
    int? horizontalDpi,
    int? verticalDpi,
    int? copies,
    bool? usePrinterDefaults,
  }) {
    return PageSetup(
      orientation: orientation ?? this.orientation,
      paperSize: paperSize ?? this.paperSize,
      paperWidth: paperWidth ?? this.paperWidth,
      paperHeight: paperHeight ?? this.paperHeight,
      scale: scale ?? this.scale,
      fitToPage: fitToPage ?? this.fitToPage,
      fitToWidth: fitToWidth ?? this.fitToWidth,
      fitToHeight: fitToHeight ?? this.fitToHeight,
      firstPageNumber: firstPageNumber ?? this.firstPageNumber,
      useFirstPageNumber: useFirstPageNumber ?? this.useFirstPageNumber,
      pageOrder: pageOrder ?? this.pageOrder,
      blackAndWhite: blackAndWhite ?? this.blackAndWhite,
      draft: draft ?? this.draft,
      cellComments: cellComments ?? this.cellComments,
      errors: errors ?? this.errors,
      horizontalDpi: horizontalDpi ?? this.horizontalDpi,
      verticalDpi: verticalDpi ?? this.verticalDpi,
      copies: copies ?? this.copies,
      usePrinterDefaults: usePrinterDefaults ?? this.usePrinterDefaults,
    );
  }

  /// Serializes this [PageSetup] to an OpenXML `<pageSetup .../>` element.
  ///
  /// [relationshipId] keeps the `r:id` link to an existing printer settings
  /// part when re-saving a file that had one. Returns an empty string when
  /// there is nothing to write.
  String toXmlString({String? relationshipId}) {
    if (!hasPageSetupAttributes && relationshipId == null) return '';

    final sb = StringBuffer('<pageSetup');
    if (paperSize != null) sb.write(' paperSize="${paperSize!.code}"');
    if (paperHeight != null) {
      sb.write(' paperHeight="${_escapeXml(paperHeight!)}"');
    }
    if (paperWidth != null) {
      sb.write(' paperWidth="${_escapeXml(paperWidth!)}"');
    }
    if (scale != null) sb.write(' scale="${scale!.clamp(10, 400)}"');
    if (firstPageNumber != null) {
      sb.write(' firstPageNumber="$firstPageNumber"');
    }
    if (fitToWidth != null) sb.write(' fitToWidth="$fitToWidth"');
    if (fitToHeight != null) sb.write(' fitToHeight="$fitToHeight"');
    if (pageOrder != null) sb.write(' pageOrder="${pageOrder!.xmlValue}"');
    if (orientation != null) {
      sb.write(' orientation="${orientation!.xmlValue}"');
    }
    if (usePrinterDefaults != null) {
      sb.write(' usePrinterDefaults="${usePrinterDefaults! ? 1 : 0}"');
    }
    if (blackAndWhite != null) {
      sb.write(' blackAndWhite="${blackAndWhite! ? 1 : 0}"');
    }
    if (draft != null) sb.write(' draft="${draft! ? 1 : 0}"');
    if (cellComments != null) {
      sb.write(' cellComments="${cellComments!.xmlValue}"');
    }
    if (useFirstPageNumber != null) {
      sb.write(' useFirstPageNumber="${useFirstPageNumber! ? 1 : 0}"');
    }
    if (errors != null) sb.write(' errors="${errors!.xmlValue}"');
    if (horizontalDpi != null) sb.write(' horizontalDpi="$horizontalDpi"');
    if (verticalDpi != null) sb.write(' verticalDpi="$verticalDpi"');
    if (copies != null) sb.write(' copies="$copies"');
    if (relationshipId != null) sb.write(' r:id="$relationshipId"');
    sb.write('/>');
    return sb.toString();
  }

  @override
  List<Object?> get props => [
        orientation,
        paperSize,
        paperWidth,
        paperHeight,
        scale,
        fitToPage,
        fitToWidth,
        fitToHeight,
        firstPageNumber,
        useFirstPageNumber,
        pageOrder,
        blackAndWhite,
        draft,
        cellComments,
        errors,
        horizontalDpi,
        verticalDpi,
        copies,
        usePrinterDefaults,
      ];
}

/// Represents the OpenXML `<pageMargins>` element. All values are in inches.
class PageMargins extends Equatable {
  final double left;
  final double right;
  final double top;
  final double bottom;

  /// Distance from the top edge of the page to the header.
  final double header;

  /// Distance from the bottom edge of the page to the footer.
  final double footer;

  /// Defaults match Excel's "Normal" margins.
  const PageMargins({
    this.left = 0.7,
    this.right = 0.7,
    this.top = 0.75,
    this.bottom = 0.75,
    this.header = 0.3,
    this.footer = 0.3,
  });

  /// Excel's "Normal" margin preset.
  static const normal = PageMargins();

  /// Excel's "Wide" margin preset.
  static const wide = PageMargins(
      left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5);

  /// Excel's "Narrow" margin preset.
  static const narrow = PageMargins(left: 0.25, right: 0.25);

  /// Same [margin] on all four sides.
  const PageMargins.all(double margin,
      {this.header = 0.3, this.footer = 0.3})
      : left = margin,
        right = margin,
        top = margin,
        bottom = margin;

  /// Creates margins from centimetres (converted to inches for storage).
  factory PageMargins.fromCentimeters({
    double left = 1.78,
    double right = 1.78,
    double top = 1.91,
    double bottom = 1.91,
    double header = 0.76,
    double footer = 0.76,
  }) {
    double toInches(double cm) => cm / 2.54;
    return PageMargins(
      left: toInches(left),
      right: toInches(right),
      top: toInches(top),
      bottom: toInches(bottom),
      header: toInches(header),
      footer: toInches(footer),
    );
  }

  PageMargins copyWith({
    double? left,
    double? right,
    double? top,
    double? bottom,
    double? header,
    double? footer,
  }) {
    return PageMargins(
      left: left ?? this.left,
      right: right ?? this.right,
      top: top ?? this.top,
      bottom: bottom ?? this.bottom,
      header: header ?? this.header,
      footer: footer ?? this.footer,
    );
  }

  /// Serializes to an OpenXML `<pageMargins .../>` element.
  String toXmlString() {
    String fmt(double v) {
      final s = v.toString();
      return s.endsWith('.0') ? s.substring(0, s.length - 2) : s;
    }

    return '<pageMargins left="${fmt(left)}" right="${fmt(right)}"'
        ' top="${fmt(top)}" bottom="${fmt(bottom)}"'
        ' header="${fmt(header)}" footer="${fmt(footer)}"/>';
  }

  @override
  List<Object?> get props => [left, right, top, bottom, header, footer];
}

/// Represents the OpenXML `<printOptions>` element.
class PrintOptions extends Equatable {
  /// Print cell gridlines.
  final bool? gridLines;

  /// Print row and column headings (A, B, C / 1, 2, 3).
  final bool? headings;

  /// Center the printed content horizontally on the page.
  final bool? horizontalCentered;

  /// Center the printed content vertically on the page.
  final bool? verticalCentered;

  /// Legacy flag used together with [gridLines]; preserved on round-trip.
  final bool? gridLinesSet;

  const PrintOptions({
    this.gridLines,
    this.headings,
    this.horizontalCentered,
    this.verticalCentered,
    this.gridLinesSet,
  });

  bool get isEmpty =>
      gridLines == null &&
      headings == null &&
      horizontalCentered == null &&
      verticalCentered == null &&
      gridLinesSet == null;

  PrintOptions copyWith({
    bool? gridLines,
    bool? headings,
    bool? horizontalCentered,
    bool? verticalCentered,
    bool? gridLinesSet,
  }) {
    return PrintOptions(
      gridLines: gridLines ?? this.gridLines,
      headings: headings ?? this.headings,
      horizontalCentered: horizontalCentered ?? this.horizontalCentered,
      verticalCentered: verticalCentered ?? this.verticalCentered,
      gridLinesSet: gridLinesSet ?? this.gridLinesSet,
    );
  }

  /// Serializes to an OpenXML `<printOptions .../>` element, or an empty
  /// string when no option is set.
  String toXmlString() {
    if (isEmpty) return '';
    final sb = StringBuffer('<printOptions');
    if (horizontalCentered != null) {
      sb.write(' horizontalCentered="${horizontalCentered! ? 1 : 0}"');
    }
    if (verticalCentered != null) {
      sb.write(' verticalCentered="${verticalCentered! ? 1 : 0}"');
    }
    if (headings != null) sb.write(' headings="${headings! ? 1 : 0}"');
    if (gridLines != null) sb.write(' gridLines="${gridLines! ? 1 : 0}"');
    if (gridLinesSet != null) {
      sb.write(' gridLinesSet="${gridLinesSet! ? 1 : 0}"');
    }
    sb.write('/>');
    return sb.toString();
  }

  @override
  List<Object?> get props =>
      [gridLines, headings, horizontalCentered, verticalCentered, gridLinesSet];
}
