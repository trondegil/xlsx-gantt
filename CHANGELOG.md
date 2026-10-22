# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.1] - 2026-10-22

### Changed
- Logos are bigger: `GanttStyle.logo_scale` now defaults to `1.0`, so the
  logo fills the height of the header rows instead of 70 % of it. Set
  `logo_scale=0.7` to get the old size.
- The README example's logo is now a kringle, and the finished workbook
  is included as `examples/bakery-plan.xlsx`.
- PyPI page: new summary and keywords, plus Changelog and Examples
  links. The README's test-suite list covers the CLI, week numbering and
  examples.

### Fixed
- README `LICENSE` links are absolute, so they work on PyPI.
- Contributing guide named the wrong default branch (`main`, not `master`).

---

## [0.4.0] - 2026-10-22

### Added
- `--color RRGGBB` CLI option builds the chart's theme from one hex
  colour (`GanttTheme.from_color`), so brand colours no longer need
  Python. It can't be combined with `--theme`.
- README section *AI-Assisted Planning*: a prompt to paste into any AI
  assistant, the plan it produced (`examples/bakery-plan.json`), and the
  resulting chart, built with `--color` and a sample logo
  (`examples/bakery-logo.png`).

---

## [0.3.0] - 2026-10-22

### Added
- `GanttStyle.week_start` sets the first day of the week for the row-1
  week bands (`0` = Monday … `6` = Sunday, as in `datetime.weekday()`).
  Monday keeps ISO 8601 week numbers; any other start day counts the week
  containing 1 January as week 1, so a week spanning New Year stays in
  one band. Default `0`, so existing charts are unchanged.
- `--week-start` CLI option taking a weekday name, e.g.
  `xlsx-gantt plan.json --week-start sunday`.

### Changed
- README Quick Start leads with the JSON file and the one-line CLI
  command; the full Python dict example follows it.
- New `assets/social-preview.png` (1280×640) for the GitHub social card.

---

## [0.2.3] - 2026-07-27

### Changed
- Trove classifier moved from `Development Status :: 3 - Alpha` to
  `4 - Beta`: the feature set is complete for the documented scope, the
  public API has been stable across 0.2.x (the one rename in 0.2.0 kept
  deprecated aliases), and the suite runs 115 tests with ruff and mypy
  clean in CI.

---

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
