import 'dart:convert';
import 'dart:io';

void main() {
  final file = File(r'C:\Users\gonoj\.gemini\antigravity-ide\brain\bea1754d-82fb-4240-885d-ee4da0cee60a\.system_generated\logs\transcript.jsonl');
  if (!file.existsSync()) {
    print('File does not exist');
    return;
  }

  final lines = file.readAsLinesSync();
  final targetFiles = <String>{};

  for (final line in lines) {
    try {
      final json = jsonDecode(line);
      final toolCalls = json['tool_calls'] as List?;
      if (toolCalls != null) {
        for (final tc in toolCalls) {
          final args = tc['args'] as Map?;
          if (args != null) {
            final target = args['TargetFile'] as String?;
            if (target != null) {
              targetFiles.add(target);
            }
          }
        }
      }
    } catch (_) {}
  }

  print('Files modified in the transcript:');
  for (final tf in targetFiles) {
    print('- $tf');
  }
}
