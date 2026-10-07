// excel_community vs excel_plus vs the original excel package: writing
// (cell().value and appendRow), encoding and reading, each run in a fresh
// AOT-compiled process (like a Flutter release build), the libraries
// alternating run by run. Prints a Markdown table with the median of 3 runs.
//
//   dart run compare.dart [rows...]      default: 10000 100000
//
// The original excel package needs older archive and xml versions, so it
// lives in its own package (original/) with the same runner.
import 'dart:io';

const runs = 3;
const limit = Duration(minutes: 5);

Map<String, int> parse(String output) {
  final line = output
      .split('\n')
      .firstWhere(
        (l) => l.startsWith('RESULT:'),
        orElse: () => throw StateError('no result in:\n$output'),
      );
  return {
    for (final part in line.substring(7).trim().split('|'))
      part.split(':')[0]: int.parse(part.split(':')[1]),
  };
}

/// One run, or `null` when it takes longer than [limit].
Future<Map<String, int>?> runOnce(String exe, List<String> args) async {
  final process = await Process.start(exe, args);
  final stdout = process.stdout
      .transform(const SystemEncoding().decoder)
      .join();
  final stderr = process.stderr
      .transform(const SystemEncoding().decoder)
      .join();
  final code = await process.exitCode.timeout(
    limit,
    onTimeout: () {
      process.kill();
      return -1;
    },
  );
  if (code == -1) return null;
  if (code != 0) throw StateError('${args.join(' ')} failed:\n${await stderr}');
  return parse(await stdout);
}

/// The median of each value, or `null` when a run timed out.
Map<String, int>? median(List<Map<String, int>?> results) {
  if (results.isEmpty || results.contains(null)) return null;
  int pick(String key) =>
      (results.map((r) => r![key]!).toList()..sort())[results.length ~/ 2];
  return {for (final key in results.first!.keys) key: pick(key)};
}

Future<String> compile(String script, String dir, String exe) async {
  final get = await Process.run('dart', ['pub', 'get'], workingDirectory: dir);
  if (get.exitCode != 0)
    throw StateError('dart pub get in $dir failed:\n${get.stderr}');
  final r = await Process.run('dart', [
    'compile',
    'exe',
    script,
    '-o',
    exe,
  ], workingDirectory: dir);
  if (r.exitCode != 0)
    throw StateError('compile $dir/$script failed:\n${r.stderr}');
  return exe;
}

String seconds(int ms) => '${(ms / 1000).toStringAsFixed(2)} s';
String total(Map<String, int>? r) => r == null
    ? '> ${limit.inMinutes} min'
    : seconds(r['build']! + r['encode']!);

Future<void> main(List<String> args) async {
  final counts = args.isEmpty ? [10000, 100000] : args.map(int.parse).toList();
  final tmp = Directory.systemTemp.createTempSync('excel_bench_');
  final ext = Platform.isWindows ? '.exe' : '';
  final sep = Platform.pathSeparator;
  final current = await compile(
    'compare_single.dart',
    '.',
    '${tmp.path}${sep}current$ext',
  );
  final original = await compile(
    'compare_single.dart',
    'original',
    '${tmp.path}${sep}original$ext',
  );
  final libraries = {
    'excel_community': (current, 'community'),
    'excel_plus': (current, 'plus'),
    'excel (original)': (original, 'original'),
  };

  print(
    'Dart ${Platform.version.split(' ').first}, AOT, 10 mixed columns, median of $runs runs\n',
  );
  print(
    '| Rows | Library | cell().value + encode | appendRow + encode | Read | File size |',
  );
  print('| ---: | :--- | ---: | ---: | ---: | ---: |');
  for (final rows in counts) {
    // Every library reads the same file, written by excel_community.
    final file = '${tmp.path}${sep}data_$rows.xlsx';
    await Process.run(current, ['community', 'append', '$rows', file]);
    // The libraries alternate run by run, so a slower moment of the machine
    // does not land on a single library.
    final results = {
      for (final name in libraries.keys)
        name: {
          'cell': <Map<String, int>?>[],
          'append': <Map<String, int>?>[],
          'read': <Map<String, int>?>[],
        },
    };
    for (var run = 0; run < runs; run++) {
      for (final MapEntry(key: name, value: (exe, id)) in libraries.entries) {
        final r = results[name]!;
        // After a timeout, the remaining runs of that case are skipped.
        if (!r['cell']!.contains(null))
          r['cell']!.add(await runOnce(exe, [id, 'cell', '$rows']));
        if (!r['append']!.contains(null))
          r['append']!.add(await runOnce(exe, [id, 'append', '$rows']));
        if (!r['read']!.contains(null))
          r['read']!.add(await runOnce(exe, [id, 'read', file]));
      }
    }
    for (final name in libraries.keys) {
      final cell = median(results[name]!['cell']!);
      final append = median(results[name]!['append']!);
      final read = median(results[name]!['read']!);
      if (read != null && read['cells']! < rows * 10) {
        throw StateError('$name read only ${read['cells']} cells');
      }
      final size = (cell ?? append)?['size'];
      print(
        '| $rows | $name | ${total(cell)} | ${total(append)} '
        '| ${read == null ? '> ${limit.inMinutes} min' : seconds(read['read']!)} '
        '| ${size == null ? '-' : '${(size / 1024).round()} KB'} |',
      );
    }
  }
  tmp.deleteSync(recursive: true);
}
