"""
xlsx_gantt/cli.py
~~~~~~~~~~~~~~~~~~
Command-line interface: build a Gantt chart .xlsx from a JSON file.

Usage::

    xlsx-gantt input.json -o gantt.xlsx [--theme ocean] [--logo logo.png]
                                        [--week-start sunday]

The JSON schema is documented on :meth:`~xlsx_gantt.GanttChart.from_json`,
which this CLI is a thin wrapper around.
"""
from __future__ import annotations

import argparse
import sys

from .chart import GanttChart
from .style import GanttStyle
from .themes import GanttTheme

# Indexed like datetime.weekday() / GanttStyle.week_start.
_WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday",
             "friday", "saturday", "sunday"]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="xlsx-gantt",
        description="Generate a styled Gantt chart .xlsx from a JSON file.",
    )
    parser.add_argument("input", help="JSON file describing the chart")
    parser.add_argument(
        "-o", "--output", default="gantt.xlsx",
        help="output .xlsx path (default: gantt.xlsx)",
    )
    parser.add_argument(
        "--theme", choices=GanttTheme.list_themes(), default=None,
        help="named colour theme",
    )
    parser.add_argument("--logo", default=None, help="path to a logo image")
    parser.add_argument(
        "--week-start", choices=_WEEKDAYS, default="monday",
        help="first day of the week for the week bands (default: monday)",
    )
    args = parser.parse_args(argv)

    style = GanttTheme.get(args.theme) if args.theme else GanttStyle()
    style.week_start = _WEEKDAYS.index(args.week_start)

    try:
        chart = GanttChart.from_json(
            args.input,
            style     = style,
            logo_path = args.logo,
        )
    except OSError as err:
        print(f"xlsx-gantt: error: cannot read '{args.input}': {err}",
              file=sys.stderr)
        return 1
    except ValueError as err:
        print(f"xlsx-gantt: error: {err}", file=sys.stderr)
        return 1

    try:
        chart.generate_excel(args.output)
    except OSError as err:
        print(f"xlsx-gantt: error: cannot write '{args.output}': {err}",
              file=sys.stderr)
        return 1

    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
