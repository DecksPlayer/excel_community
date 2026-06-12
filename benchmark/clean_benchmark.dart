import 'dart:convert';
import 'dart:io';

class Result {
  final int build;
  final int encode;
  final int total;
  final int size;

  Result(this.build, this.encode, this.total, this.size);

  factory Result.parse(String output) {
    // Expected output format: RESULT:build:123|encode:456|total:579|size:1000
    final line = output.split('\n').firstWhere((l) => l.startsWith('RESULT:'));
    final parts = line.substring(7).split('|');
    int build = 0;
    int encode = 0;
    int total = 0;
    int size = 0;
    for (var part in parts) {
      final kv = part.split(':');
      if (kv[0] == 'build') build = int.parse(kv[1]);
      if (kv[0] == 'encode') encode = int.parse(kv[1]);
      if (kv[0] == 'total') total = int.parse(kv[1]);
      if (kv[0] == 'size') size = int.parse(kv[1]);
    }
    return Result(build, encode, total, size);
  }
}

Future<Result> runProcess(String library, int cellsCount) async {
  final isOriginal = library == 'original';
  final workingDir = isOriginal ? '../.benchmark/excel_baseline' : '.';
  final script = isOriginal ? 'benchmark_single.dart' : 'run_single.dart';
  final args = isOriginal
      ? [cellsCount.toString()]
      : [library, cellsCount.toString()];

  final process = await Process.start('dart', [
    'run',
    script,
    ...args,
  ], workingDirectory: workingDir);

  final output = await process.stdout.transform(utf8.decoder).join();
  final stderrOutput = await process.stderr.transform(utf8.decoder).join();

  if (process.exitCode != 0 && !output.contains('RESULT:')) {
    print('Error running process: $stderrOutput');
    exit(1);
  }

  return Result.parse(output);
}

void printTable(String title, Result community, Result plus, Result original) {
  print('\n=== $title ===');
  print(
    '-----------------------------------------------------------------------------',
  );
  print(
    '${'Library'.padRight(20)} | ${'Build (ms)'.padLeft(12)} | ${'Encode (ms)'.padLeft(12)} | ${'Total (ms)'.padLeft(12)} | ${'Speed Ratio'.padLeft(12)}',
  );
  print(
    '-----------------------------------------------------------------------------',
  );

  final double commTotal = community.total.toDouble();
  final double plusTotal = plus.total.toDouble();
  final double origTotal = original.total.toDouble();
  final double minTotal = commTotal < plusTotal
      ? (commTotal < origTotal ? commTotal : origTotal)
      : (plusTotal < origTotal ? plusTotal : origTotal);

  final commRatio = commTotal / minTotal;
  final plusRatio = plusTotal / minTotal;
  final origRatio = origTotal / minTotal;

  print(
    '${'excel_community'.padRight(20)} | ${community.build.toString().padLeft(12)} | ${community.encode.toString().padLeft(12)} | ${community.total.toString().padLeft(12)} | ${commRatio.toStringAsFixed(2).padLeft(11)}x',
  );
  print(
    '${'excel_plus'.padRight(20)} | ${plus.build.toString().padLeft(12)} | ${plus.encode.toString().padLeft(12)} | ${plus.total.toString().padLeft(12)} | ${plusRatio.toStringAsFixed(2).padLeft(11)}x',
  );
  print(
    '${'excel_original'.padRight(20)} | ${original.build.toString().padLeft(12)} | ${original.encode.toString().padLeft(12)} | ${original.total.toString().padLeft(12)} | ${origRatio.toStringAsFixed(2).padLeft(11)}x',
  );
  print(
    '-----------------------------------------------------------------------------',
  );
  print(
    'File sizes: excel_community = ${community.size} bytes, excel_plus = ${plus.size} bytes, excel_original = ${original.size} bytes\n',
  );
}

Future<void> main() async {
  print(
    '=============================================================================',
  );
  print('EXCEL LIBRARIES CLEAN COLD-START SCALING BENCHMARK');
  print(
    'Running each benchmark run in a completely fresh Dart VM process (from scratch)',
  );
  print(
    '=============================================================================',
  );

  print('\nRunning Benchmark 1: 10K Cells...');
  final comm10k = await runProcess('community', 10000);
  final plus10k = await runProcess('plus', 10000);
  final orig10k = await runProcess('original', 10000);
  printTable(
    '10K Cells Benchmark (1,000 rows x 10 cols)',
    comm10k,
    plus10k,
    orig10k,
  );

  print('\nRunning Benchmark 2: 100K Cells...');
  final comm100k = await runProcess('community', 100000);
  final plus100k = await runProcess('plus', 100000);
  final orig100k = await runProcess('original', 100000);
  printTable(
    '100K Cells Benchmark (10,000 rows x 10 cols)',
    comm100k,
    plus100k,
    orig100k,
  );

  print('\nRunning Benchmark 3: 5M Cells...');
  final comm5m = await runProcess('community', 5000000);
  final plus5m = await runProcess('plus', 5000000);

  print(
    'Running original excel package on 5M cells (this usually times out)...',
  );
  Result orig5m;
  try {
    orig5m = await runProcess(
      'original',
      5000000,
    ).timeout(const Duration(minutes: 2));
  } catch (e) {
    orig5m = Result(-1, -1, -1, -1);
  }

  print('\n=== 5M Cells Benchmark (500,000 rows x 10 cols) ===');
  print(
    '-----------------------------------------------------------------------------',
  );
  print(
    '${'Library'.padRight(20)} | ${'Build (ms)'.padLeft(12)} | ${'Encode (ms)'.padLeft(12)} | ${'Total (ms)'.padLeft(12)} | ${'Speed Ratio'.padLeft(12)}',
  );
  print(
    '-----------------------------------------------------------------------------',
  );
  final double commTotal = comm5m.total.toDouble();
  final double plusTotal = plus5m.total.toDouble();
  final double minTotal = commTotal < plusTotal ? commTotal : plusTotal;
  final commRatio = commTotal / minTotal;
  final plusRatio = plusTotal / minTotal;
  print(
    '${'excel_community'.padRight(20)} | ${comm5m.build.toString().padLeft(12)} | ${comm5m.encode.toString().padLeft(12)} | ${comm5m.total.toString().padLeft(12)} | ${commRatio.toStringAsFixed(2).padLeft(11)}x',
  );
  print(
    '${'excel_plus'.padRight(20)} | ${plus5m.build.toString().padLeft(12)} | ${plus5m.encode.toString().padLeft(12)} | ${plus5m.total.toString().padLeft(12)} | ${plusRatio.toStringAsFixed(2).padLeft(11)}x',
  );
  if (orig5m.total == -1) {
    print(
      '${'excel_original'.padRight(20)} | ${'Timeout'.padLeft(12)} | ${'Timeout'.padLeft(12)} | ${'Timeout (>2m)'.padLeft(12)} | ${'—'.padLeft(12)}',
    );
  } else {
    final origRatio = orig5m.total / minTotal;
    print(
      '${'excel_original'.padRight(20)} | ${orig5m.build.toString().padLeft(12)} | ${orig5m.encode.toString().padLeft(12)} | ${orig5m.total.toString().padLeft(12)} | ${origRatio.toStringAsFixed(2).padLeft(11)}x',
    );
  }
  print(
    '-----------------------------------------------------------------------------',
  );
  print(
    'File sizes: excel_community = ${comm5m.size} bytes, excel_plus = ${plus5m.size} bytes, excel_original = ${orig5m.size == -1 ? "Timeout" : orig5m.size} bytes\n',
  );
}
