<p align="center">
  <img src="https://raw.githubusercontent.com/trondegil/xlsx-gantt/master/icon.png" alt="xlsx-gantt icon" width="96"/>
</p>

<h1 align="center">xlsx-gantt</h1>

<p align="center">
  <a href="https://pypi.org/project/xlsx-gantt/"><img src="https://img.shields.io/pypi/v/xlsx-gantt?color=blue" alt="PyPI version"/></a>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white" alt="Python 3.10+"/></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License"/></a>
  <a href="https://pypi.org/project/openpyxl/"><img src="https://img.shields.io/badge/openpyxl-%E2%89%A53.1-informational?logo=microsoft-excel&logoColor=white" alt="openpyxl ≥ 3.1"/></a>
  <a href="https://github.com/trondegil/xlsx-gantt"><img src="https://img.shields.io/badge/github-trondegil%2Fxlsx--gantt-181717?logo=github&logoColor=white" alt="GitHub repository"/></a>
</p>

<p align="center">
  Turn Python data into polished, presentation-ready Gantt charts — no Excel required, no GUI needed.<br/>
  Built on <a href="https://openpyxl.readthedocs.io/">openpyxl</a>. Zero heavy dependencies.
</p>

<p align="center">
  <img src="https://raw.githubusercontent.com/trondegil/xlsx-gantt/master/screenshot.png" alt="xlsx-gantt screenshot" width="100%"/>
</p>

---

## Features

- **Polished charts out of the box** — pick one of eight ready-made colour themes, or generate a matching palette from your brand colour
- **No Excel needed to create them** — charts are plain `.xlsx` files anyone can open, filter, and edit in Excel
- **Simple to use** — describe your project as sections and tasks, get a presentation-ready chart; even directly from the command line: `xlsx-gantt chart.json -o gantt.xlsx`
- **Shows what matters** — progress bars per task, milestones, who's responsible for what, and automatic time-estimate totals
- **Your language, your logo** — day and week labels are fully configurable, and your logo can sit in the chart header
- **Lightweight** — just [openpyxl](https://openpyxl.readthedocs.io/) and [Pillow](https://pypi.org/project/Pillow/); no pandas, no matplotlib, no Java bridge

---

## Installation

```bash
pip install xlsx-gantt
```

This includes everything — all logo image formats (PNG, JPEG, BMP, GIF, TIFF)
work out of the box.

### Development install

```bash
git clone https://github.com/trondegil/xlsx-gantt.git
cd xlsx-gantt
pip install -e ".[dev]"
```

---

## Quick Start

```python
from datetime import datetime
from xlsx_gantt import GanttChart, GanttTheme

sections = [
    {
        "name": "Design",
        "tasks": [
            {
                "name": "Requirements gathering",
                "time_estimate": 3,
                "progress": 100,
                "ranges": [
                    {"start": datetime(2026, 1, 5), "end": datetime(2026, 1, 9), "color": "0070C0"},
                ],
                "annotations": {"Alice": "R", "Bob": "S"},
            },
            {
                "name": "Architecture review",
                "time_estimate": 2,
                "progress": 80,
                "ranges": [
                    {"start": datetime(2026, 1, 12), "end": datetime(2026, 1, 14), "color": "0070C0"},
                ],
                "annotations": {"Alice": "R", "Bob": "R"},
            },
        ],
    },
    {
        "name": "Development",
        "tasks": [
            {
                "name": "Backend",
                "time_estimate": 10,
                "progress": 50,
                "ranges": [
                    {"start": datetime(2026, 1, 19), "end": datetime(2026, 2, 6), "color": "00B050"},
                ],
                "annotations": {"Alice": "S", "Bob": "R"},
            },
            {
                "name": "Frontend",
                "time_estimate": 8,
                "progress": 20,
                "ranges": [
                    {"start": datetime(2026, 1, 26), "end": datetime(2026, 2, 13), "color": "00B050"},
                ],
                "annotations": {"Alice": "R", "Bob": "S"},
            },
        ],
    },
    {
        "name": "Milestones",
        "tasks": [
            {
                "name": "Beta release",
                "time_estimate": None,
                "progress": None,
                "ranges": [
                    {"start": datetime(2026, 2, 9), "end": datetime(2026, 2, 9), "color": "FF0000"},
                ],
            },
        ],
    },
]

style = GanttTheme.get("ocean")

chart = GanttChart(
    sections=sections,
    start_date=datetime(2026, 1, 5),    # Monday
    end_date=datetime(2026, 2, 15),     # Sunday
    project_name="My Project",
    resource_names=["Alice", "Bob"],
    style=style,
    logo_path="logo.png",               # optional
)

# Save to disk
chart.generate_excel("gantt_chart.xlsx")

# — or — get raw bytes (no file written)
xlsx_bytes = chart.generate_excel_bytes()
print("Done!")
```

### From a JSON file

Keep the project plan in a JSON file instead of Python code and load it in
one line:

```python
from xlsx_gantt import GanttChart, GanttTheme

chart = GanttChart.from_json("chart.json", style=GanttTheme.get("ocean"))
chart.generate_excel("gantt.xlsx")
```

The JSON mirrors the dict API, with dates as ISO strings — see
[Command-Line Usage](#command-line-usage) for a full example file. The same
file also works without any Python at all: `xlsx-gantt chart.json -o gantt.xlsx`.

---

## Typed Dataclass API (optional)

Instead of plain dicts you can use the typed `Section`, `Task`, and `DateRange`
dataclasses.  Both forms are fully interchangeable and may be freely mixed in the
same `sections` list.

```python
from datetime import datetime
from xlsx_gantt import GanttChart, GanttTheme, Section, Task, DateRange

sections = [
    Section(
        name="Design",
        tasks=[
            Task(
                name="Requirements gathering",
                time_estimate=3,
                progress=100,
                ranges=[
                    DateRange(
                        start=datetime(2026, 1, 5),
                        end=datetime(2026, 1, 9),
                        color="0070C0",
                    )
                ],
                annotations={"Alice": "R", "Bob": "S"},
            ),
        ],
    ),
]

chart = GanttChart(
    sections=sections,
    start_date=datetime(2026, 1, 5),
    end_date=datetime(2026, 2, 15),
    project_name="My Project",
    resource_names=["Alice", "Bob"],
    style=GanttTheme.get("ocean"),
)
chart.generate_excel("gantt_chart.xlsx")
```

### Mixing dicts and dataclasses

```python
from xlsx_gantt import Section, Task

# Dict section alongside dataclass section — both work
sections = [
    {"name": "Phase 1", "tasks": [{"name": "Kickoff", "progress": 100}]},
    Section(name="Phase 2", tasks=[Task(name="Development", time_estimate=10)]),
]
```

### to_dict / from_dict

Each dataclass provides `to_dict()` (serialize back to the dict API) and
`from_dict()` (construct from a raw dict, returning `None` for invalid input):

```python
from xlsx_gantt import Section, Task, DateRange
from datetime import datetime

dr = DateRange(start=datetime(2026, 1, 5), end=datetime(2026, 1, 9), color="0070C0")
print(dr.to_dict())
# {'start': datetime(2026, 1, 5, 0, 0), 'end': datetime(2026, 1, 9, 0, 0), 'color': '0070C0'}

task = Task.from_dict({"name": "Research", "time_estimate": 3, "progress": 80})
print(task.name)  # Research
```

---

## In-Memory Output

`generate_excel_bytes()` returns the raw `.xlsx` bytes without writing to disk.
Useful for streaming from a web server or attaching to an e-mail:

```python
chart = GanttChart(sections=sections, ...)

# Flask / Django example
xlsx_bytes = chart.generate_excel_bytes()

response = HttpResponse(
    xlsx_bytes,
    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
response["Content-Disposition"] = 'attachment; filename="gantt.xlsx"'
```

---

## Command-Line Usage

Installing the package also installs an `xlsx-gantt` command that builds a
chart from a JSON file — no Python required:

```bash
xlsx-gantt chart.json -o gantt.xlsx --theme ocean [--logo logo.png]
```

The JSON mirrors the dict API, with dates as ISO strings (`YYYY-MM-DD`):

```json
{
  "project_name": "Website Redesign",
  "start_date": "2026-03-02",
  "end_date": "2026-04-12",
  "resource_names": ["Alice", "Bob"],
  "sections": [
    {
      "name": "Design",
      "tasks": [
        {
          "name": "Research",
          "time_estimate": 3,
          "progress": 80,
          "ranges": [
            {"start": "2026-03-02", "end": "2026-03-06", "color": "0070C0"}
          ],
          "annotations": {"Alice": "R", "Bob": "S"}
        }
      ]
    }
  ]
}
```

`--theme` accepts any name from the [Themes](#themes) table. Invalid input is
reported on stderr with exit code 1.

The CLI is a thin wrapper around `GanttChart.from_json()`, so the same JSON
file can be loaded from Python too:

```python
chart = GanttChart.from_json("chart.json")
```

---

## Data Structure

### Sections (dict form)

```python
{
    "name": str,       # section label (merged across all its task rows in column A)
    "tasks": [ ... ]   # list of task dicts (see below)
}
```

### Tasks (dict form)

```python
{
    "name": str,                      # task label
    "time_estimate": float | None,    # hours / days / any unit; summed in the Total row
    "progress": float | None,         # 0–100 %; rendered as a solid DataBar
    "ranges": [                       # one or more coloured bars on the timeline
        {
            "start": datetime,        # inclusive start date
            "end":   datetime,        # inclusive end date (start == end → milestone)
            "color": "RRGGBB",        # 6-digit hex, no '#'; falls back to theme bar colour
        },
        ...
    ],
    "annotations": {                  # optional resource role markers
        "Alice": "R",                 # R = Responsible
        "Bob":   "S",                 # S = Support
    },
}
```

> **Note:** `time_estimate`, `progress`, `ranges`, and `annotations` are all optional.
> Omitting them or passing `None` is handled gracefully.

---

## Themes

Eight built-in themes are available:

| Name | Description | Base colour |
|------|-------------|-------------|
| `amber` | Warm corporate orange | `FFC000` |
| `ocean` | Deep professional blue | `1A5276` |
| `forest` | Environmental green | `1E8449` |
| `crimson` | Bold, high-impact red | `922B21` |
| `slate` | Minimal neutral gray | `5D6D7E` |
| `royal_purple` | Original default palette | `7030A0` |
| `midnight` | Modern near-black | `1B2631` |
| `teal` | Fresh blue-green | `148F77` |

```python
from xlsx_gantt import GanttTheme

# By static method
style = GanttTheme.ocean()

# By name (case-insensitive; spaces and hyphens treated as underscores)
style = GanttTheme.get("royal_purple")

# Generated from any base hex colour
style = GanttTheme.from_color("#2E86C1")

# List all available theme names
print(GanttTheme.list_themes())
```

---

## Custom Styles

Pass a `GanttStyle` dataclass instance to override any visual property:

```python
from xlsx_gantt import GanttChart, GanttStyle

style = GanttStyle(
    header_bg        = "2C3E50",
    header_fg        = "FFFFFF",
    bar_color        = "E74C3C",
    row_bg_even      = "F2F3F4",
    section_name_bg  = "D5D8DC",
    annotation_r_bg  = "C8FFCC",   # green tint for "R: Responsible"
    annotation_s_bg  = "FFE4B5",   # amber tint for "S: Support"
    col_label_bg     = "FFC000",
    col_label_fg     = "000000",
)

chart = GanttChart(sections=..., style=style, ...)
```

### Key `GanttStyle` fields

| Field | Default | Description |
|-------|---------|-------------|
| `header_bg` | `"7030A0"` | Background for header rows (rows 1–2) |
| `header_fg` | `"FFFFFF"` | Text colour for header rows |
| `bar_color` | `"9966CC"` | Default Gantt bar fill |
| `progress_bar_color` | `"7030A0"` | Solid DataBar fill colour |
| `weekend_bg` | `"EFEFEF"` | Column fill for Sat/Sun |
| `row_bg_even` | `None` | Even data-row background (`None` = white) |
| `row_bg_odd` | `None` | Odd data-row background (`None` = white) |
| `section_name_bg` | `None` | First row of each section |
| `col_label_bg` | `"FFC000"` | Row 3 — Activity / Task / date numbers |
| `annotation_r_bg` | `None` | Background for "R" annotation cells (`annotation_a_bg` still accepted) |
| `annotation_s_bg` | `None` | Background for "S" annotation cells |
| `total_row_bg` | `None` | Total row background (`None` → `header_bg`) |
| `day_names` | `("Mon", …, "Sun")` | Day abbreviations for row 2 (Mon-first) — set your own for other locales |
| `week_label_format` | `"Week {week}"` | Row-1 week band label; `{week}` and `{year}` placeholders |
| `logo_scale` | `0.7` | Logo size as a fraction of the A1:B2 area |
| `font_name` | `"Calibri"` | Font used throughout the sheet |
| `col_width_task` | `25.0` | Task column width (Excel units) |
| `col_width_date` | `3.0` | Width of each day column |

All fields and their defaults are documented in the `GanttStyle` dataclass docstring.

---

## Colour Utilities

The package also exports the colour helpers used internally by the theme engine:

```python
from xlsx_gantt import contrast_text, darken, lighten, rotate_hue

# Pick white or black for maximum WCAG contrast against a background
text = contrast_text("1A5276")   # → "FFFFFF"

# Blend a colour toward black or white
darker  = darken("FFC000", 0.4)
lighter = lighten("FFC000", 0.6)

# Rotate the hue by a given number of degrees (0–360)
comp = rotate_hue("1A5276", 180)
```

---

## For Power Users

- **In-memory output** — `generate_excel_bytes()` returns raw `.xlsx` bytes without touching the file system; drop it straight into a Flask/Django response, e-mail it, or cache it in Redis
- **Theme engine** — `GanttTheme.from_color("#2E86C1")` derives a complete WCAG-contrast-checked palette from any hex colour; the colour helpers `darken`, `lighten`, `rotate_hue`, and `contrast_text` are exported for your own theme logic
- **Typed dataclass API** — `Section`, `Task`, and `DateRange` dataclasses give you IDE auto-completion and type checking; plain dicts still work and can be freely mixed in the same list
- **Multiple bars per row** — paint several independent date ranges on a single task row, each with its own hex colour; any range where `start == end` renders as a milestone
- **Solid progress DataBars** — gradient-free Excel 2010 DataBar conditional formatting, injected via `x14` extension XML post-processing
- **Fine-grained styling** — every colour, font, column width, and row height is a `GanttStyle` field; section names merge vertically, headers freeze, and label columns get auto-filters
- **Locale control** — `GanttStyle.day_names` and `week_label_format` (with `{week}`/`{year}` placeholders) replace the default English labels
- **Production-safe error handling** — malformed sections, tasks, ranges, and annotations are silently skipped; no unhandled exceptions in live environments

---

## Project Structure

```
xlsx-gantt/
├── xlsx_gantt/                # Library package
│   ├── __init__.py            # Public re-exports
│   ├── style.py               # GanttStyle configuration dataclass
│   ├── models.py              # Section, Task, DateRange input dataclasses
│   ├── chart.py               # GanttChart builder
│   ├── themes.py              # GanttTheme + colour utilities
│   ├── cli.py                 # xlsx-gantt command-line interface
│   ├── py.typed               # PEP 561 marker
│   └── _xlsx_patch.py         # Internal: solid DataBar XML post-processor
├── tests/
│   ├── __init__.py
│   ├── test_malformed_input.py
│   ├── test_new_features.py
│   ├── test_readme_example.py
│   └── test_v020.py
├── pyproject.toml
├── requirements.txt
├── requirements-dev.txt
├── CHANGELOG.md
├── LICENSE
└── README.md
```

---

## Running the Tests

```bash
pip install -r requirements-dev.txt
pytest
```

The test suite covers:
- Valid input (dict API and dataclass API)
- In-memory output (`generate_excel_bytes`)
- Mixed dict + dataclass inputs
- `to_dict` / `from_dict` round-trips
- `patch_solid_databars` with both file paths and `BytesIO`
- Malformed sections, tasks, ranges, and annotations
- Edge-case chart configurations (empty date ranges, `None` fields, reversed dates, etc.)

---

## Requirements

| Package | Version | Notes |
|---------|---------|-------|
| Python | ≥ 3.10 | |
| [openpyxl](https://pypi.org/project/openpyxl/) | ≥ 3.1.0 | Required |
| [Pillow](https://pypi.org/project/Pillow/) | ≥ 10.0.0 | Required |

---

## Contributing

Contributions are welcome! Here is the recommended workflow:

1. **Fork** the repository and create a feature branch:

   ```bash
   git checkout -b feature/my-improvement
   ```

2. **Install** the project in editable mode with dev dependencies:

   ```bash
   pip install -e ".[dev]"
   ```

3. **Make your changes.** Keep the following guidelines in mind:
   - Follow the existing code style (PEP 8, type hints, docstrings).
   - Add or update tests in `tests/` for any changed behaviour.
   - All colour values must be 6-digit hex strings without a leading `#`.

4. **Run the test suite** and make sure all tests pass:

   ```bash
   pytest
   ```

5. **Commit** with a clear, descriptive message:

   ```bash
   git commit -m "feat: add X / fix Y"
   ```

6. **Open a pull request** against the `main` branch describing what you changed and why.

### Reporting issues

Please open a GitHub Issue and include:
- Python version (`python --version`)
- openpyxl version (`pip show openpyxl`)
- A minimal reproducible example or the full traceback.

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for the full text.

> **Dependency licences**
> - [openpyxl](https://pypi.org/project/openpyxl/) — MIT
> - [Pillow](https://pypi.org/project/Pillow/) — HPND (historically permissive, MIT-compatible)
