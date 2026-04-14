"""
tests/test_readme_example.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Ensures the Quick Start example shown in README.md runs without error
and produces a valid .xlsx file.

The code here is kept intentionally in sync with the README example.
If you change the Quick Start snippet in README.md, update this test
to match.
"""
from __future__ import annotations

import os
import tempfile
import zipfile
from datetime import datetime

import pytest

from xlsx_gantt import GanttChart, GanttTheme


# ── README Quick Start data ───────────────────────────────────────────────────

SECTIONS = [
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


# ── Helper ────────────────────────────────────────────────────────────────────

def _build_chart() -> GanttChart:
    """Construct the chart exactly as shown in the README Quick Start."""
    style = GanttTheme.get("ocean")
    return GanttChart(
        sections=SECTIONS,
        start_date=datetime(2026, 1, 5),   # Monday
        end_date=datetime(2026, 2, 15),    # Sunday
        project_name="My Project",
        resource_names=["Alice", "Bob"],
        style=style,
        logo_path="logo.png",              # optional – silently skipped if absent
    )


# ── Tests ─────────────────────────────────────────────────────────────────────

class TestReadmeExample:
    """The README Quick Start example runs and produces a valid xlsx."""

    def test_generates_without_error(self, tmp_path):
        """generate_excel() completes without raising an exception."""
        chart = _build_chart()
        chart.generate_excel(str(tmp_path / "gantt_chart.xlsx"))

    def test_output_file_is_created(self, tmp_path):
        """The output .xlsx file exists after generation."""
        out = tmp_path / "gantt_chart.xlsx"
        _build_chart().generate_excel(str(out))
        assert out.exists(), "Output file was not created"

    def test_output_is_valid_zip(self, tmp_path):
        """A valid .xlsx is a ZIP archive — verify the file is not corrupted."""
        out = tmp_path / "gantt_chart.xlsx"
        _build_chart().generate_excel(str(out))
        assert zipfile.is_zipfile(out), "Output is not a valid ZIP / xlsx file"

    def test_output_contains_worksheet(self, tmp_path):
        """The archive must contain the primary worksheet XML."""
        out = tmp_path / "gantt_chart.xlsx"
        _build_chart().generate_excel(str(out))
        with zipfile.ZipFile(out) as z:
            assert "xl/worksheets/sheet1.xml" in z.namelist()

    def test_output_is_non_empty(self, tmp_path):
        """The output file should have a reasonable size (> 5 KB)."""
        out = tmp_path / "gantt_chart.xlsx"
        _build_chart().generate_excel(str(out))
        assert out.stat().st_size > 5_000, "Output file is suspiciously small"

    def test_all_themes_generate_without_error(self, tmp_path):
        """Every built-in theme should work with the README example data."""
        from xlsx_gantt import GanttTheme

        for theme_name in GanttTheme.list_themes():
            style = GanttTheme.get(theme_name)
            chart = GanttChart(
                sections=SECTIONS,
                start_date=datetime(2026, 1, 5),
                end_date=datetime(2026, 2, 15),
                project_name="My Project",
                resource_names=["Alice", "Bob"],
                style=style,
            )
            out = tmp_path / f"gantt_{theme_name}.xlsx"
            chart.generate_excel(str(out))  # must not raise
