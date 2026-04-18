# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
## [1.0.8] - 2026-04-07
- add funding text
## [1.0.7] - 2026-04-07

- Fix underline preservation bug in CellStyle constructor and parsing logic.
- Add strikethrough text style support.
- Add comprehensive tests for underline and strikethrough preservation.
- Update Flutter example with underline and strikethrough demos.
## [1.0.6] - 2026-03-17
- Fix color accessibility.
## [1.0.5] - 2026-03-17
- Apply clean code fix minors bugs.
## [1.0.4] - 2026-03-15
- Apply Clean Architecture in Save class.
- Refactor Sheet class.
- Refactor Color class.
- Refactor Border class.
- Refactor Formula class.
- Refactor NumberFormat class.

## [1.0.3] - 2026-03-15

- Renamed main example to `example/main.dart` for better visibility on pub.dev.
- Significantly increased documentation coverage (>28%) across the public API.
- Added detailed documentation for all chart types and cell value classes.

## [1.0.2] - 2026-03-15

## [1.0.1] - 2026-03-14

- Reorganized tests: moved manual debug scripts to `test/manual/`.
- Fixed static analysis errors in `parse.dart` and `save_file.dart`.
- Improved package metadata for pub.dev.

## [1.0.0] - 2026-03-14

- Initial Release of the community fork.
- Forked from [excel](https://github.com/justkawal/excel)
- Updated minimum Dart SDK to 3.6.0
- Migrate from `dart:html` to `package:web`
- Update packages to latest versions
- Added support for charts

