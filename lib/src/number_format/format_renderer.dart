part of '../../excel_community.dart';

// Best-effort renderer for ECMA-376 §18.8.30 number format codes.
//
// Covers the ~50 built-in standard formats and the common custom patterns:
// digit templates (`000-000-0000`), thousands grouping and scaling
// (`#,##0,"K"`), fractions (`# ?/?`, `# ?/8`), scientific and engineering
// notation, currency/locale tags (`[$€-2]`, `[$-404]`) and date/time codes
// including era years, CJK AM/PM designators and elapsed time. Conditional
// sections (`[>100]`) and locale-specific month names are not implemented.

CellValue? _resolveDisplayValue(CellValue? value) {
  if (value is FormulaCellValue) return value.cachedValue;
  return value;
}

// ---------------------------------------------------------------------------
// Numeric rendering
// ---------------------------------------------------------------------------

/// Renders [value] as text following a numeric format code such as
/// `"#,##0.00"`, `"0%"` or `"$#,##0.00_);($#,##0.00)"`.
String renderNumericValue(String formatCode, num value) {
  if (formatCode.isEmpty || formatCode == 'General' || formatCode == '@') {
    return _renderGeneralNumber(value);
  }

  final sections = _splitFormatSections(formatCode);
  if (value < 0 && sections.length >= 2) {
    return _renderNumericSection(sections[1], value.abs());
  }
  if (value == 0 && sections.length >= 3 && sections[2].isNotEmpty) {
    return _renderNumericSection(sections[2], 0);
  }
  final rendered = _renderNumericSection(sections[0], value.abs());
  // Like Excel, no minus sign when the value displays as zero (-0.001 with
  // "0.00").
  if (value < 0 && rendered != _renderNumericSection(sections[0], 0)) {
    return '-$rendered';
  }
  return rendered;
}

String _renderGeneralNumber(num value) {
  if (value is int) return value.toString();
  final d = value.toDouble();
  if (d == d.roundToDouble() && d.abs() < 1e15) {
    return d.toInt().toString();
  }
  var s = d.toStringAsPrecision(10);
  if (!s.contains('e') && !s.contains('E') && s.contains('.')) {
    s = s.replaceFirst(RegExp(r'0+$'), '');
    s = s.replaceFirst(RegExp(r'\.$'), '');
  }
  return s;
}

List<String> _splitFormatSections(String formatCode) {
  final sections = <String>[];
  final buffer = StringBuffer();
  var inQuotes = false;
  var i = 0;
  while (i < formatCode.length) {
    final c = formatCode[i];
    if (c == '"') {
      inQuotes = !inQuotes;
      buffer.write(c);
      i++;
    } else if (!inQuotes && c == '[') {
      final end = formatCode.indexOf(']', i + 1);
      final stop = end == -1 ? formatCode.length : end + 1;
      buffer.write(formatCode.substring(i, stop));
      i = stop;
    } else if (!inQuotes && c == ';') {
      sections.add(buffer.toString());
      buffer.clear();
      i++;
    } else {
      buffer.write(c);
      i++;
    }
  }
  sections.add(buffer.toString());
  return sections;
}

String _renderNumericSection(String section, num absValue) {
  if (section.trim().isEmpty) return '';

  if (RegExp('E[+-]', caseSensitive: false)
      .hasMatch(_stripQuotedAndBracketed(section))) {
    return _renderScientific(section, absValue);
  }
  return _renderFraction(section, absValue) ??
      _renderDigitTemplate(section, absValue);
}

// A numeric section is rendered as a template: digit placeholders (`0`, `#`,
// `?`) receive the digits of the value and every other token keeps its
// position, so codes like `000-000-0000` or `"Total: "#,##0` work.

enum _NumTokenKind { literal, digit, point, comma, percent }

class _NumToken {
  final _NumTokenKind kind;
  final String text;

  const _NumToken(this.kind, this.text);
}

List<_NumToken> _tokenizeNumericSection(String section) {
  final tokens = <_NumToken>[];
  var hasPoint = false;
  var i = 0;
  while (i < section.length) {
    final c = section[i];
    if (c == '"') {
      final end = section.indexOf('"', i + 1);
      final stop = end == -1 ? section.length : end;
      tokens.add(_NumToken(_NumTokenKind.literal, section.substring(i + 1, stop)));
      i = end == -1 ? section.length : end + 1;
    } else if (c == '\\' && i + 1 < section.length) {
      tokens.add(_NumToken(_NumTokenKind.literal, section[i + 1]));
      i += 2;
    } else if (c == '[') {
      final end = section.indexOf(']', i + 1);
      if (end == -1) break;
      final content = section.substring(i + 1, end);
      // [$€-2] is a currency symbol; colors and conditions render nothing.
      if (content.startsWith(r'$')) {
        final dashIdx = content.indexOf('-');
        tokens.add(_NumToken(_NumTokenKind.literal,
            dashIdx == -1 ? content.substring(1) : content.substring(1, dashIdx)));
      }
      i = end + 1;
    } else if ((c == '_' || c == '*') && i + 1 < section.length) {
      // Alignment fill/spacer characters have no meaningful rendering in
      // plain display text.
      i += 2;
    } else if (c == '0' || c == '#' || c == '?') {
      tokens.add(_NumToken(_NumTokenKind.digit, c));
      i++;
    } else if (c == '.' && !hasPoint) {
      hasPoint = true;
      tokens.add(const _NumToken(_NumTokenKind.point, '.'));
      i++;
    } else if (c == ',') {
      tokens.add(const _NumToken(_NumTokenKind.comma, ','));
      i++;
    } else if (c == '%') {
      tokens.add(const _NumToken(_NumTokenKind.percent, '%'));
      i++;
    } else {
      tokens.add(_NumToken(_NumTokenKind.literal, c));
      i++;
    }
  }
  return tokens;
}

String _renderDigitTemplate(String section, num absValue) {
  final tokens = _tokenizeNumericSection(section);
  final pointIdx = tokens.indexWhere((t) => t.kind == _NumTokenKind.point);
  final integerEnd = pointIdx == -1 ? tokens.length : pointIdx;

  final integerSlots = <int>[];
  final decimalSlots = <int>[];
  for (var i = 0; i < tokens.length; i++) {
    if (tokens[i].kind != _NumTokenKind.digit) continue;
    (i < integerEnd ? integerSlots : decimalSlots).add(i);
  }
  if (integerSlots.isEmpty && decimalSlots.isEmpty) {
    return tokens
        .map((t) => t.kind == _NumTokenKind.point ? '.' : t.text)
        .join();
  }

  // A comma between integer placeholders turns on thousands grouping; a
  // comma right after the number scales the value down by 1000.
  var grouping = false;
  var scaleCommas = 0;
  final silentCommas = <int>{};
  for (var i = 0; i < tokens.length; i++) {
    if (tokens[i].kind != _NumTokenKind.comma) continue;
    var prev = i - 1;
    while (prev >= 0 && tokens[prev].kind == _NumTokenKind.comma) {
      prev--;
    }
    var next = i + 1;
    while (next < tokens.length && tokens[next].kind == _NumTokenKind.comma) {
      next++;
    }
    final afterNumber = prev >= 0 &&
        (tokens[prev].kind == _NumTokenKind.digit ||
            tokens[prev].kind == _NumTokenKind.point);
    if (!afterNumber) continue; // literal comma
    silentCommas.add(i);
    if (next < integerEnd && tokens[next].kind == _NumTokenKind.digit) {
      grouping = true;
    } else {
      scaleCommas++;
    }
  }

  var value = absValue.toDouble();
  final percents = tokens.where((t) => t.kind == _NumTokenKind.percent).length;
  for (var i = 0; i < percents; i++) {
    value *= 100;
  }
  for (var i = 0; i < scaleCommas; i++) {
    value /= 1000;
  }

  final decimals = decimalSlots.length;
  final factor = pow(10, decimals);
  final fixed = ((value * factor).round() / factor).toStringAsFixed(decimals);
  final parts = fixed.split('.');
  final integerDigits = parts[0] == '0' ? '' : parts[0];
  final decimalDigits = decimals > 0 ? parts[1] : '';

  final out = List<String>.generate(tokens.length, (i) {
    final t = tokens[i];
    if (t.kind == _NumTokenKind.comma) {
      return silentCommas.contains(i) ? '' : ',';
    }
    return t.kind == _NumTokenKind.digit ? '' : t.text;
  });

  // Integer digits, right-aligned into the placeholders; surplus digits go
  // into the first one.
  if (integerSlots.isEmpty) {
    if (integerDigits.isNotEmpty && pointIdx != -1) {
      out[pointIdx] = '$integerDigits.';
    }
  } else {
    final n = integerSlots.length;
    final rendered = <String>[];
    for (var k = 0; k < n; k++) {
      final d = integerDigits.length - n + k;
      if (d >= 0) {
        rendered.add(k == 0
            ? integerDigits.substring(0, d + 1)
            : integerDigits[d]);
      } else {
        rendered.add(_emptyPlaceholder(tokens[integerSlots[k]].text));
      }
    }
    final hasInnerLiterals = tokens
        .sublist(integerSlots.first, integerSlots.last + 1)
        .any((t) => t.kind == _NumTokenKind.literal);
    if (grouping && !hasInnerLiterals) {
      final joined = rendered.join();
      final digitsStart = joined.indexOf(RegExp(r'\d'));
      out[integerSlots.first] = digitsStart == -1
          ? joined
          : joined.substring(0, digitsStart) +
              _groupThousands(joined.substring(digitsStart));
    } else {
      for (var k = 0; k < n; k++) {
        out[integerSlots[k]] = rendered[k];
      }
    }
  }

  // Decimal digits, left-aligned; trailing zeros are dropped for `#` and
  // shown as spaces for `?`.
  for (var k = 0; k < decimals; k++) {
    final digit = decimalDigits[k];
    final placeholder = tokens[decimalSlots[k]].text;
    final restIsZero = decimalDigits.substring(k).replaceAll('0', '').isEmpty;
    out[decimalSlots[k]] = (placeholder != '0' && restIsZero)
        ? _emptyPlaceholder(placeholder)
        : digit;
  }

  return out.join();
}

/// What an unused digit placeholder displays.
String _emptyPlaceholder(String placeholder) => switch (placeholder) {
      '0' => '0',
      '?' => ' ',
      _ => '',
    };

/// Renders fraction codes such as `# ?/?`, `# ??/??`, `?/?` or `# ?/8`.
/// Returns `null` when [section] is not a fraction format.
String? _renderFraction(String section, num absValue) {
  if (!_stripQuotedAndBracketed(section).contains('/')) return null;
  final match = RegExp(r'^(.*?)(?:([#0?]+)(\s+))?([#0?]+)/([#0?]+|[1-9][0-9]*)(.*)$')
      .firstMatch(section);
  if (match == null) return null;

  final prefix = _renderLiterals(match.group(1)!);
  final integerMask = match.group(2);
  final separator = match.group(3) ?? '';
  final numeratorMask = match.group(4)!;
  final denominatorSpec = match.group(5)!;
  final suffix = _renderLiterals(match.group(6)!);

  final hasInteger = integerMask != null;
  var integerPart = hasInteger ? absValue.floor() : 0;
  final fraction = absValue - integerPart;

  var numerator = 0;
  var denominator = 1;
  final fixedDenominator = int.tryParse(denominatorSpec);
  if (fixedDenominator != null) {
    denominator = fixedDenominator;
    numerator = (fraction * denominator).round();
  } else {
    // Last continued-fraction convergent whose denominator fits the
    // placeholders (Excel does not use other "best" approximations: 0.3 with
    // one digit is 1/3, not 2/7).
    final maxDenominator = pow(10, denominatorSpec.length).toInt() - 1;
    var hPrev = 1, hPrev2 = 0, kPrev = 0, kPrev2 = 1;
    var x = fraction.toDouble();
    for (var i = 0; i < 64; i++) {
      final a = x.floor();
      final h = a * hPrev + hPrev2;
      final k = a * kPrev + kPrev2;
      if (k > maxDenominator) break;
      numerator = h;
      denominator = k;
      hPrev2 = hPrev;
      hPrev = h;
      kPrev2 = kPrev;
      kPrev = k;
      final rest = x - a;
      if (rest < 1e-10) break;
      x = 1 / rest;
    }
  }
  if (hasInteger && numerator == denominator) {
    integerPart += 1;
    numerator = 0;
  }

  var integerText = '';
  if (hasInteger) {
    integerText = integerPart == 0
        ? _emptyPlaceholder(integerMask[integerMask.length - 1])
        : integerPart.toString();
  }

  final fractionWidth = numeratorMask.length + 1 + denominatorSpec.length;
  if (hasInteger && numerator == 0) {
    // Excel blanks the fraction part when it is zero.
    final shown = integerText.trim().isEmpty ? '0' : integerText;
    return '$prefix$shown$separator${' ' * fractionWidth}$suffix';
  }

  final numeratorText = numeratorMask.contains('0')
      ? numerator.toString().padLeft(numeratorMask.length, '0')
      : numeratorMask.contains('?')
          ? numerator.toString().padLeft(numeratorMask.length)
          : numerator.toString();
  final denominatorText = fixedDenominator != null
      ? denominatorSpec
      : denominatorSpec.contains('?')
          ? denominator.toString().padRight(denominatorSpec.length)
          : denominator.toString();
  return '$prefix$integerText${hasInteger ? separator : ''}'
      '$numeratorText/$denominatorText$suffix';
}

/// Renders scientific (`0.00E+00`) and engineering (`##0.0E+0`) notation.
/// With `#` in a multi-digit mantissa the exponent is a multiple of the
/// number of integer placeholders, as in Excel.
String _renderScientific(String section, num absValue) {
  final stripped = _stripQuotedAndBracketed(section);
  final match =
      RegExp(r'([#0?,]*)(?:\.([#0?]*))?E([+-])([0#?]+)', caseSensitive: false)
          .firstMatch(stripped);
  if (match == null) return _renderGeneralNumber(absValue);

  final integerMask = match.group(1)!.replaceAll(',', '');
  final decimalDigits = (match.group(2) ?? '').length;
  final expSign = match.group(3)!;
  final expDigits = match.group(4)!.length;
  final integerPlaces = max(1, integerMask.length);
  final engineering = integerMask.length > 1 && integerMask.contains('#');

  var exponent = 0;
  var mantissa = absValue.toDouble();
  if (mantissa != 0) {
    var magnitude = 0;
    var scaled = mantissa;
    while (scaled >= 10) {
      scaled /= 10;
      magnitude++;
    }
    while (scaled < 1) {
      scaled *= 10;
      magnitude--;
    }
    final step = engineering ? integerPlaces : 1;
    exponent = engineering
        ? (magnitude / integerPlaces).floor() * integerPlaces
        : magnitude - (integerPlaces - 1);
    mantissa = absValue / pow(10, exponent);
    // Rounding can carry into a new digit (9.999 -> 10.00); move the
    // exponent instead.
    if (double.parse(mantissa.toStringAsFixed(decimalDigits)) >=
        pow(10, integerPlaces)) {
      exponent += step;
      mantissa = absValue / pow(10, exponent);
    }
  }

  // Excel fills every integer placeholder for zero ("##0.0E+0" -> 000.0E+0).
  final mantissaStr = absValue == 0
      ? '${'0' * integerPlaces}${decimalDigits > 0 ? '.${'0' * decimalDigits}' : ''}'
      : mantissa.toStringAsFixed(decimalDigits);
  final expStr = exponent.abs().toString().padLeft(expDigits, '0');
  final expSignStr = exponent < 0 ? '-' : (expSign == '+' ? '+' : '');
  return '${mantissaStr}E$expSignStr$expStr';
}

String _groupThousands(String digits) {
  final reversed = digits.split('').reversed.toList();
  final buffer = StringBuffer();
  for (var i = 0; i < reversed.length; i++) {
    if (i != 0 && i % 3 == 0) buffer.write(',');
    buffer.write(reversed[i]);
  }
  return buffer.toString().split('').reversed.join();
}

String _stripQuotedAndBracketed(String s) {
  final buffer = StringBuffer();
  var inQuotes = false;
  var i = 0;
  while (i < s.length) {
    final c = s[i];
    if (c == '"') {
      inQuotes = !inQuotes;
      i++;
    } else if (inQuotes) {
      i++;
    } else if (c == '[') {
      final end = s.indexOf(']', i + 1);
      i = end == -1 ? s.length : end + 1;
    } else {
      buffer.write(c);
      i++;
    }
  }
  return buffer.toString();
}

String _renderLiterals(String s) {
  final buffer = StringBuffer();
  var i = 0;
  while (i < s.length) {
    final c = s[i];
    if (c == '"') {
      final end = s.indexOf('"', i + 1);
      if (end == -1) {
        buffer.write(s.substring(i + 1));
        break;
      }
      buffer.write(s.substring(i + 1, end));
      i = end + 1;
    } else if (c == '\\' && i + 1 < s.length) {
      buffer.write(s[i + 1]);
      i += 2;
    } else if (c == '[') {
      final end = s.indexOf(']', i + 1);
      if (end == -1) {
        i = s.length;
        continue;
      }
      final content = s.substring(i + 1, end);
      if (content.startsWith(r'$')) {
        final dashIdx = content.indexOf('-');
        buffer.write(dashIdx == -1
            ? content.substring(1)
            : content.substring(1, dashIdx));
      }
      i = end + 1;
    } else if ((c == '_' || c == '*') && i + 1 < s.length) {
      // Alignment fill/spacer characters have no meaningful rendering in
      // plain display text.
      i += 2;
    } else {
      buffer.write(c);
      i++;
    }
  }
  return buffer.toString();
}

// ---------------------------------------------------------------------------
// Date / time rendering
// ---------------------------------------------------------------------------

const _monthNames = [
  'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December', //
];
const _monthNamesShort = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov',
  'Dec', //
];
const _weekdayNames = [
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
  'Sunday', //
];
const _weekdayNamesShort = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

class _DtToken {
  // literal, y, era, month, d, h, minute, s, subsec, ampm
  final String type;
  final int length;
  final String? text;
  final bool elapsed;

  const _DtToken(
      {required this.type, this.length = 0, this.text, this.elapsed = false});

  const _DtToken.literal(this.text)
      : type = 'literal',
        length = 0,
        elapsed = false;

  _DtToken withType(String newType) =>
      _DtToken(type: newType, length: length, text: text, elapsed: elapsed);
}

/// Renders a date and/or time as text following a date/time format code
/// such as `"mm-dd-yy"`, `"h:mm AM/PM"` or `"[h]:mm:ss"`.
///
/// Elapsed-time masks (`[h]`, `[mm]`, `[ss]`) count from Excel's day zero:
/// [elapsedDays] whole days plus [hour]/[minute]/[second]. An [hour] greater
/// than 23 is also rendered as-is under `[h]`.
///
/// Seconds are rounded to the precision the code shows (`ss`, `ss.0`, ...),
/// carrying into minutes, hours and days like Excel.
///
/// A locale tag such as `[$-411]` selects the AM/PM designators and, for the
/// Japanese calendar (`[$-411]` or calendar type 03, e.g. `[$-30411]`), the
/// era shown by `e` (era year) and `g`/`gg`/`ggg` (era name). Other
/// calendars show the Gregorian year for `e`, as Excel does.
String renderDateTimeValue(
  String formatCode, {
  int? year,
  int? month,
  int? day,
  required int hour,
  required int minute,
  required int second,
  int millisecond = 0,
  int elapsedDays = 0,
}) {
  final (tokens, lcid) = _tokenizeDateTimeFormat(formatCode);

  // Round to the shown precision: whole seconds, or tenths/hundredths/
  // thousandths with ss.0 / ss.00 / ss.000.
  final digits = tokens
      .where((t) => t.type == 'subsec')
      .fold<int>(0, (m, t) => max(m, min(t.length, 3)));
  final unit = pow(10, 3 - digits).toInt();
  var ms = (millisecond / unit).round() * unit;
  var s = second, mi = minute, h = hour, days = elapsedDays;
  int? y = year, mo = month, d = day;
  if (ms >= 1000) {
    ms -= 1000;
    s++;
  }
  if (s >= 60) {
    s -= 60;
    mi++;
  }
  if (mi >= 60) {
    mi -= 60;
    h++;
  }
  if (h >= 24 && y != null && mo != null && d != null) {
    h -= 24;
    days++;
    final next = DateTime.utc(y, mo, d).add(const Duration(days: 1));
    y = next.year;
    mo = next.month;
    d = next.day;
  }

  return _renderDateTimeTokens(_resolveMinuteVsMonth(tokens),
      lcid: lcid,
      year: y,
      month: mo,
      day: d,
      hour: h,
      minute: mi,
      second: s,
      millisecond: ms,
      elapsedDays: days);
}

(List<_DtToken>, int?) _tokenizeDateTimeFormat(String formatCode) {
  const runChars = {'y', 'e', 'g', 'm', 'd', 'h', 's'};
  final tokens = <_DtToken>[];
  int? lcid;
  var i = 0;
  while (i < formatCode.length) {
    final c = formatCode[i];
    if (c == '"') {
      final end = formatCode.indexOf('"', i + 1);
      final stop = end == -1 ? formatCode.length : end;
      tokens.add(_DtToken.literal(formatCode.substring(i + 1, stop)));
      i = end == -1 ? formatCode.length : end + 1;
      continue;
    }
    if (c == '\\' && i + 1 < formatCode.length) {
      tokens.add(_DtToken.literal(formatCode[i + 1]));
      i += 2;
      continue;
    }
    if (c == '[') {
      final end = formatCode.indexOf(']', i + 1);
      if (end == -1) {
        i = formatCode.length;
        continue;
      }
      final content = formatCode.substring(i + 1, end);
      if (RegExp(r'^[hH]+$').hasMatch(content)) {
        tokens.add(_DtToken(type: 'h', length: content.length, elapsed: true));
      } else if (RegExp(r'^[mM]+$').hasMatch(content)) {
        tokens.add(
            _DtToken(type: 'minute', length: content.length, elapsed: true));
      } else if (RegExp(r'^[sS]+$').hasMatch(content)) {
        tokens.add(_DtToken(type: 's', length: content.length, elapsed: true));
      } else {
        // [$-404] / [$€-2]: the hex part after '-' is the locale id. Colors
        // and conditions have no rendering effect.
        final locale = RegExp(r'^\$[^-]*-([0-9A-Fa-f]+)$').firstMatch(content);
        if (locale != null) {
          // Low 16 bits: language; next byte: calendar type.
          lcid = int.parse(locale.group(1)!, radix: 16);
        }
      }
      i = end + 1;
      continue;
    }
    final upperRest = formatCode.substring(i).toUpperCase();
    if (upperRest.startsWith('AM/PM')) {
      tokens.add(const _DtToken(type: 'ampm', length: 5));
      i += 5;
      continue;
    }
    if (upperRest.startsWith('A/P')) {
      tokens.add(const _DtToken(type: 'ampm', length: 3));
      i += 3;
      continue;
    }
    // Chinese and Japanese AM/PM designators.
    if (formatCode.startsWith('上午/下午', i)) {
      tokens.add(const _DtToken(type: 'ampm', length: 5, text: 'zh'));
      i += 5;
      continue;
    }
    if (formatCode.startsWith('午前/午後', i)) {
      tokens.add(const _DtToken(type: 'ampm', length: 5, text: 'ja'));
      i += 5;
      continue;
    }
    // Fractions of a second right after the seconds: ss.0, ss.00, ss.000
    if (c == '.' &&
        tokens.isNotEmpty &&
        tokens.last.type == 's' &&
        i + 1 < formatCode.length &&
        formatCode[i + 1] == '0') {
      var j = i + 1;
      while (j < formatCode.length && formatCode[j] == '0') {
        j++;
      }
      tokens.add(_DtToken(type: 'subsec', length: j - i - 1));
      i = j;
      continue;
    }
    final lower = c.toLowerCase();
    if (runChars.contains(lower)) {
      var j = i;
      while (j < formatCode.length && formatCode[j].toLowerCase() == lower) {
        j++;
      }
      tokens.add(_DtToken(
          type: switch (lower) { 'e' => 'era', 'g' => 'eraName', _ => lower },
          length: j - i));
      i = j;
      continue;
    }
    tokens.add(_DtToken.literal(c));
    i++;
  }
  return (tokens, lcid);
}

List<_DtToken> _resolveMinuteVsMonth(List<_DtToken> tokens) {
  for (var i = 0; i < tokens.length; i++) {
    final t = tokens[i];
    if (t.type != 'm' || t.elapsed) continue;

    String? prevType;
    for (var j = i - 1; j >= 0; j--) {
      if (tokens[j].type != 'literal') {
        prevType = tokens[j].type;
        break;
      }
    }
    String? nextType;
    for (var j = i + 1; j < tokens.length; j++) {
      if (tokens[j].type != 'literal') {
        nextType = tokens[j].type;
        break;
      }
    }
    tokens[i] =
        t.withType((prevType == 'h' || nextType == 's') ? 'minute' : 'month');
  }
  return tokens;
}

/// Whether [lcid] uses the Japanese emperor-era calendar: the ja-JP
/// language or calendar type 03.
bool _isJapaneseEraCalendar(int? lcid) =>
    lcid != null && ((lcid & 0xFFFF) == 0x411 || ((lcid >> 16) & 0xFF) == 0x03);

/// Japanese eras: start date, first year, initial, name.
const _japaneseEras = [
  (2019, 5, 1, 2019, 'R', '令和'),
  (1989, 1, 8, 1989, 'H', '平成'),
  (1926, 12, 25, 1926, 'S', '昭和'),
  (1912, 7, 30, 1912, 'T', '大正'),
  (1868, 1, 1, 1868, 'M', '明治'),
];

(int, String, String)? _japaneseEra(int year, int month, int day) {
  final date = DateTime.utc(year, month, day);
  for (final (y, m, d, first, initial, name) in _japaneseEras) {
    if (!date.isBefore(DateTime.utc(y, m, d))) return (year - first + 1, initial, name);
  }
  return null;
}

/// Year shown by the `e` code: the Japanese era year under the Japanese
/// calendar, otherwise the Gregorian year (as Excel shows it).
int _eraYear(int year, int month, int day, int? lcid) {
  if (!_isJapaneseEraCalendar(lcid)) return year;
  return _japaneseEra(year, month, day)?.$1 ?? year;
}

String _ampmDesignator(_DtToken token, bool isAm, int? lcid) {
  final style = token.text ??
      switch (lcid) {
        // Language part of the locale id.
        _ when lcid == null => null,
        _ when const {0x404, 0x804, 0xC04, 0x1004, 0x1404}.contains(lcid & 0xFFFF) => 'zh',
        _ when (lcid & 0xFFFF) == 0x411 => 'ja',
        _ when (lcid & 0xFFFF) == 0x412 => 'ko',
        _ => null,
      };
  return switch (style) {
    'zh' => isAm ? '上午' : '下午',
    'ja' => isAm ? '午前' : '午後',
    'ko' => isAm ? '오전' : '오후',
    _ => token.length == 3 ? (isAm ? 'A' : 'P') : (isAm ? 'AM' : 'PM'),
  };
}

String _renderDateTimeTokens(
  List<_DtToken> tokens, {
  int? lcid,
  int? year,
  int? month,
  int? day,
  required int hour,
  required int minute,
  required int second,
  int millisecond = 0,
  int elapsedDays = 0,
}) {
  final has12Hour = tokens.any((t) => t.type == 'ampm');
  final totalHours = elapsedDays * 24 + hour;
  final totalMinutes = totalHours * 60 + minute;
  final buffer = StringBuffer();
  for (final t in tokens) {
    switch (t.type) {
      case 'literal':
        buffer.write(t.text ?? '');
      case 'y':
        final y = year ?? 1900;
        buffer.write(t.length >= 3
            ? y.toString().padLeft(4, '0')
            : (y % 100).toString().padLeft(2, '0'));
      case 'era':
        final y = _eraYear(year ?? 1900, month ?? 1, day ?? 1, lcid);
        buffer.write(y.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 'month':
        final m = month ?? 1;
        if (t.length >= 4) {
          buffer.write(_monthNames[(m - 1).clamp(0, 11)]);
        } else if (t.length == 3) {
          buffer.write(_monthNamesShort[(m - 1).clamp(0, 11)]);
        } else {
          buffer.write(m.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
        }
      case 'd':
        final d = day ?? 1;
        if (t.length >= 3) {
          final weekday = DateTime.utc(year ?? 1900, month ?? 1, d).weekday - 1;
          buffer.write(t.length >= 4
              ? _weekdayNames[weekday.clamp(0, 6)]
              : _weekdayNamesShort[weekday.clamp(0, 6)]);
        } else {
          buffer.write(d.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
        }
      case 'h':
        var h = t.elapsed ? totalHours : hour;
        if (!t.elapsed && has12Hour) {
          h = h % 12;
          if (h == 0) h = 12;
        }
        buffer.write(h.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 'minute':
        final m = t.elapsed ? totalMinutes : minute;
        buffer.write(m.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 's':
        final sec = t.elapsed ? totalMinutes * 60 + second : second;
        buffer.write(sec.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 'subsec':
        // millisecond is already rounded to the shown precision.
        final shown = min(t.length, 3);
        final value = millisecond ~/ pow(10, 3 - shown).toInt();
        buffer.write('.${value.toString().padLeft(shown, '0')}${'0' * (t.length - shown)}');
      case 'eraName':
        final era = _isJapaneseEraCalendar(lcid)
            ? _japaneseEra(year ?? 1900, month ?? 1, day ?? 1)
            : null;
        if (era != null) {
          buffer.write(switch (t.length) {
            1 => era.$2,
            2 => era.$3.substring(0, 1),
            _ => era.$3,
          });
        }
      case 'ampm':
        buffer.write(_ampmDesignator(t, (hour % 24) < 12, lcid));
      default:
        break;
    }
  }
  return buffer.toString();
}
