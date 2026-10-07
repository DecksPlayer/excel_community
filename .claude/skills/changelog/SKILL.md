---
name: changelog
description: How to write the two changelogs of excel_community (Dart package and npm package). Use when adding a CHANGELOG entry, bumping a version, or preparing release notes.
---

# Skill: excel_community changelogs

The repository ships two packages, and each has its own changelog:

| File | Package | Versions |
| :--- | :--- | :--- |
| `CHANGELOG.md` (root) | Dart, `excel_community` on pub.dev | `version` in `pubspec.yaml` |
| `npm/CHANGELOG.md` | npm, `excel-community` | `version` in `npm/package.json` (its history differs from Dart's: check `npm view excel-community versions`) |

## Each changelog is different

A change goes only in the changelog of the package it affects. Do not mention npm in the Dart changelog, or Dart in the npm changelog.

**Exception: a problem solved for both.** The npm package compiles the Dart engine (`lib/`), so a fix in `lib/` that users of both packages notice goes in **both** changelogs, worded for each audience (Dart API names in one, JS API names in the other). Do not write "see the Dart changelog" in the npm one.

How to decide:

| The change is in… | Goes in |
| :--- | :--- |
| `lib/` and users of both packages notice it (a bug fix, speed, file output) | Both |
| `lib/`, but only through Dart API that npm does not expose | Dart only |
| `npm/` (`wrapper.js`, `npm/dart/*`, `index.d.ts`, `build.js`, `bench.js`, npm README) | npm only |
| `benchmark/`, `excel_flutter_example/`, Dart README | Dart only, and only if users notice it |
| A new API added to both | Both, each with its own API name (e.g. `sheet.refreshPivotTables()` in both) |

## Format

```markdown
## [2.6.0] - 2026-10-07
### Fixed
- Fixed `appendRow` getting slower as the sheet grows.

### Improved
- Improved save and read speed.

### Added
- `sheet.refreshPivotTables()` to recompute pivot tables after changing their data.

### Changed
- A `Cell` points to a position, so it no longer follows its data when rows are inserted or removed.
```

- Sections in this order: `Fixed`, `Improved`, `Added`, `Changed`; leave out the empty ones.
- One short line per change. Fixed lines start with "Fixed …" and name the problem, not the solution.
- No numbers, measurements or internals: performance figures go in the READMEs, measured separately per package (npm against exceljs, Dart against excel_plus and the original excel).
- The newest version goes first, under the file's header.

## Release

- The version in the changelog must match `pubspec.yaml` (Dart) or `npm/package.json` (npm).
- `npm/CHANGELOG.md` ships in the npm package (`files` in `npm/package.json`).
- The maintainer commits and publishes; do not commit, tag or publish unless asked.
