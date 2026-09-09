part of '../../excel_community.dart';

// Best-effort renderer for ECMA-376 §18.8.30 number format codes.
//
// This does not implement the full Excel formatting language (locale-aware
// currency symbols, exact fraction rendering, conditional threshold
// sections, etc.) but covers the ~50 built-in standard formats plus the
// most common custom currency / percentage / thousands / date-time
// patterns, which is enough to render a reasonable `displayText` for a
// cell's value.

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
  return value < 0 ? '-$rendered' : rendered;
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

  final hasPercent = _hasTopLevelChar(section, '%');
  final workingValue = hasPercent ? absValue * 100 : absValue;

  final maskMatch =
      RegExp(r'[#0][0-9#,]*(?:\.[0#]*)?|\.[0#]+').firstMatch(section);
  if (maskMatch == null) {
    return _renderLiterals(section);
  }

  final mask = maskMatch.group(0)!;
  final dotIdx = mask.indexOf('.');
  final integerMask = dotIdx == -1 ? mask : mask.substring(0, dotIdx);
  final decimalMask = dotIdx == -1 ? '' : mask.substring(dotIdx + 1);
  final decimalDigits = decimalMask.length;
  final minIntegerDigits =
      integerMask.replaceAll(',', '').replaceAll('#', '').length;
  final grouping = integerMask.contains(',');

  final numberText = _formatFixed(workingValue, decimalDigits,
      minIntegerDigits: minIntegerDigits, grouping: grouping);

  final before = section.substring(0, maskMatch.start);
  final after = section.substring(maskMatch.end);
  return '${_renderLiterals(before)}$numberText${_renderLiterals(after)}';
}

String _renderScientific(String section, num absValue) {
  final stripped = _stripQuotedAndBracketed(section);
  final match = RegExp(r'([#0]*\.?[#0]*)E([+-])([0#]+)', caseSensitive: false)
      .firstMatch(stripped);
  if (match == null) return _renderGeneralNumber(absValue);

  final mantissaMask = match.group(1)!;
  final expSign = match.group(2)!;
  final expDigits = match.group(3)!.length;
  final decimalDigits =
      mantissaMask.contains('.') ? mantissaMask.split('.').last.length : 0;

  var exponent = 0;
  var mantissa = absValue.toDouble();
  if (mantissa != 0) {
    while (mantissa >= 10) {
      mantissa /= 10;
      exponent++;
    }
    while (mantissa < 1) {
      mantissa *= 10;
      exponent--;
    }
  }

  final mantissaStr = mantissa.toStringAsFixed(decimalDigits);
  final expStr = exponent.abs().toString().padLeft(expDigits, '0');
  final expSignStr = exponent < 0 ? '-' : (expSign == '+' ? '+' : '');
  return '${mantissaStr}E$expSignStr$expStr';
}

String _formatFixed(num value, int decimalDigits,
    {required int minIntegerDigits, required bool grouping}) {
  final factor = pow(10, decimalDigits);
  final rounded = (value * factor).round() / factor;
  final fixed = rounded.toStringAsFixed(decimalDigits);
  final parts = fixed.split('.');
  var integerPart = parts[0];
  if (integerPart.length < minIntegerDigits) {
    integerPart = integerPart.padLeft(minIntegerDigits, '0');
  }
  if (grouping) {
    integerPart = _groupThousands(integerPart);
  }
  return decimalDigits > 0 ? '$integerPart.${parts[1]}' : integerPart;
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

bool _hasTopLevelChar(String s, String target) {
  var inQuotes = false;
  var i = 0;
  while (i < s.length) {
    final c = s[i];
    if (c == '"') {
      inQuotes = !inQuotes;
      i++;
    } else if (inQuotes) {
      i++;
    } else if (c == '\\') {
      i += 2;
    } else if (c == '[') {
      final end = s.indexOf(']', i + 1);
      i = end == -1 ? s.length : end + 1;
    } else if (c == target) {
      return true;
    } else {
      i++;
    }
  }
  return false;
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
  final String type; // literal, y, month, d, h, minute, s, ampm
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
/// [hour]/[minute]/[second] are treated as *totals*, so passing an [hour]
/// greater than 23 renders correctly under an elapsed-time mask (`[h]`).
String renderDateTimeValue(
  String formatCode, {
  int? year,
  int? month,
  int? day,
  required int hour,
  required int minute,
  required int second,
  int millisecond = 0,
}) {
  final tokens = _resolveMinuteVsMonth(_tokenizeDateTimeFormat(formatCode));
  return _renderDateTimeTokens(tokens,
      year: year,
      month: month,
      day: day,
      hour: hour,
      minute: minute,
      second: second,
      millisecond: millisecond);
}

List<_DtToken> _tokenizeDateTimeFormat(String formatCode) {
  const runChars = {'y', 'e', 'm', 'd', 'h', 's'};
  final tokens = <_DtToken>[];
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
      }
      // Otherwise: color/condition/locale code — no rendering effect.
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
    final lower = c.toLowerCase();
    if (runChars.contains(lower)) {
      var j = i;
      while (j < formatCode.length && formatCode[j].toLowerCase() == lower) {
        j++;
      }
      tokens.add(_DtToken(type: lower == 'e' ? 'y' : lower, length: j - i));
      i = j;
      continue;
    }
    tokens.add(_DtToken.literal(c));
    i++;
  }
  return tokens;
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

String _renderDateTimeTokens(
  List<_DtToken> tokens, {
  int? year,
  int? month,
  int? day,
  required int hour,
  required int minute,
  required int second,
  int millisecond = 0,
}) {
  final has12Hour = tokens.any((t) => t.type == 'ampm');
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
        var h = hour;
        if (!t.elapsed && has12Hour) {
          h = h % 12;
          if (h == 0) h = 12;
        }
        buffer.write(h.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 'minute':
        buffer.write(minute.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 's':
        buffer.write(second.toString().padLeft(t.length >= 2 ? 2 : 1, '0'));
      case 'ampm':
        final isAm = (hour % 24) < 12;
        buffer.write(t.length == 3 ? (isAm ? 'A' : 'P') : (isAm ? 'AM' : 'PM'));
      default:
        break;
    }
  }
  return buffer.toString();
}
