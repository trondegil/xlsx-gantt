# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.2] - 2026-07-27

Documentation-only release; the library code is unchanged from 0.2.0.

### Changed
- README notes that the dict/JSON plan format makes the package a natural
  fit for AI-assisted project planning.

---

## [0.2.1] - 2026-07-27

Documentation-only release; the library code is unchanged from 0.2.0.

### Changed
- README leads with a rendered chart and shows every built-in theme in a
  gallery under the theme table.
- README reordered for first-time readers: features, installation, quick
  start, CLI, themes, and data structure come first, with the typed
  dataclass API, in-memory output, custom styles, and colour utilities
  grouped under *Advanced Usage* and the dev instructions under
  *Development*.
- New logo, and the screenshots are regenerated so they show it.

---

## [0.2.0] - 2026-07-16

### Added
- `xlsx-gantt` command-line interface: build a chart from a JSON file
  (`xlsx-gantt input.json -o out.xlsx --theme ocean`).
- `GanttChart.from_json("chart.json")` builds a chart directly from a
  JSON file in Python (the CLI is a thin wrapper around it).
- `GanttStyle.day_names` and `GanttStyle.week_label_format` for
  locale-specific day abbreviations and week band labels
  (`{week}` and `{year}` placeholders).
- `patch_solid_databars(..., sheet=None)` patches every worksheet that
  contains a DataBar rule, not just sheet 1.
- `py.typed` marker so type checkers pick up the package's annotations
  (the `Typing :: Typed` classifier now holds).
- Ruff and mypy run in CI; `dev` extra now includes both.

### Changed
- Pillow is now a default dependency, so all logo image formats (PNG,
  JPEG, BMP, GIF, TIFF) work out of the box. The `[logo]` extra remains
  as a deprecated no-op alias.
- Style fields `annotation_a_fg` / `annotation_a_bg` renamed to
  `annotation_r_fg` / `annotation_r_bg` to match the `"R"` role they style.
  The old names are still accepted as deprecated constructor aliases.
- Logo-embedding failures now raise a `UserWarning` instead of printing
  to stdout.

### Fixed
- README images now use absolute URLs so they render on PyPI.
- LICENSE copyright named the wrong project (`excel-gantt`).
- Auto-filter range no longer overshoots into the Total row when a
  section contains invalid task entries.
- Removed a stray `asyncio_default_fixture_loop_scope` pytest option.

## [0.1.0] - 2026-04-14

### Added
- Initial release: styled Gantt chart generation via openpyxl, dict and
  dataclass APIs, eight named themes, `GanttTheme.from_color`, solid
  progress DataBars, logo embedding, in-memory output.
