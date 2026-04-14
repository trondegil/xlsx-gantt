"""
tests/test_malformed_input.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Tests that GanttChart.generate_excel() handles malformed input event data
gracefully — i.e. completes without raising an unhandled exception.

Each test calls _generate(), which builds the chart and writes it to a
temporary file (cleaned up afterwards).  If the call does not raise, the
test passes.  Where the current code already handles the case via .get()
defaults, the test acts as a regression guard.  Where the code previously
would have crashed, the corresponding defensive fix was added to
xlsx_gantt/chart.py.
"""
from __future__ import annotations

import os
import tempfile
from datetime import datetime

import pytest

from xlsx_gantt import GanttChart

# ── Shared fixtures ──────────────────────────────────────────────────────────

START = datetime(2026, 4, 6)   # Monday
END   = datetime(2026, 4, 12)  # Sunday
RESOURCES = ["Alice", "Bob"]


def _make_chart(sections, **kwargs):
    return GanttChart(
        sections=sections,
        start_date=kwargs.get("start_date", START),
        end_date=kwargs.get("end_date", END),
        project_name=kwargs.get("project_name", "Test Project"),
        resource_names=kwargs.get("resource_names", RESOURCES),
    )


def _generate(sections, **kwargs):
    """Build the chart and write to a temp .xlsx; delete afterwards."""
    chart = _make_chart(sections, **kwargs)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        fname = f.name
    try:
        chart.generate_excel(fname)
    finally:
        if os.path.exists(fname):
            os.unlink(fname)


def _valid_task(**overrides):
    base = {
        "name": "Task 1",
        "time_estimate": 5.0,
        "progress": 50,
        "ranges": [{"start": START, "end": END, "color": "0070C0"}],
        "annotations": {"Alice": "A", "Bob": "S"},
    }
    base.update(overrides)
    return base


def _valid_section(**overrides):
    base = {"name": "Section 1", "tasks": [_valid_task()]}
    base.update(overrides)
    return base


# ── Baseline ─────────────────────────────────────────────────────────────────


class TestBaseline:
    """Sanity-check that valid data still works after the defensive changes."""

    def test_valid_single_section(self):
        _generate([_valid_section()])

    def test_valid_multiple_sections(self):
        _generate([_valid_section(name="Alpha"), _valid_section(name="Beta")])

    def test_empty_sections_list(self):
        _generate([])

    def test_section_with_multiple_tasks(self):
        tasks = [_valid_task(name=f"Task {i}") for i in range(5)]
        _generate([{"name": "Multi", "tasks": tasks}])

    def test_no_resource_names(self):
        _generate([_valid_section()], resource_names=[])


# ── Malformed sections ────────────────────────────────────────────────────────


class TestMalformedSections:
    """Malformed entries at the section-dict level."""

    def test_section_missing_name_key(self):
        """Section without 'name' → falls back to empty string."""
        _generate([{"tasks": [_valid_task()]}])

    def test_section_name_is_none(self):
        """Section with name=None → None stored in cell (no crash)."""
        _generate([{"name": None, "tasks": [_valid_task()]}])

    def test_section_missing_tasks_key(self):
        """Section without 'tasks' → treated as no tasks (empty iteration)."""
        _generate([{"name": "No Tasks"}])

    def test_section_tasks_is_none(self):
        """tasks=None → treated as empty list."""
        _generate([{"name": "Null Tasks", "tasks": None}])

    def test_section_tasks_is_a_string(self):
        """tasks is a string instead of a list → treated as empty list."""
        _generate([{"name": "Bad Tasks", "tasks": "not a list"}])

    def test_section_tasks_is_a_dict(self):
        """tasks is a single dict instead of a list → treated as empty list."""
        _generate([{"name": "Dict Tasks", "tasks": {"name": "oops"}}])

    def test_section_entry_is_none(self):
        """None entry in sections list → skipped."""
        _generate([None, _valid_section()])

    def test_section_entry_is_a_string(self):
        """String entry in sections list → skipped."""
        _generate(["not a section", _valid_section()])

    def test_section_entry_is_an_integer(self):
        """Integer entry in sections list → skipped."""
        _generate([42, _valid_section()])

    def test_sections_itself_is_none(self):
        """sections=None passed to constructor → treated as empty list."""
        chart = GanttChart(
            sections=None,
            start_date=START,
            end_date=END,
        )
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            fname = f.name
        try:
            chart.generate_excel(fname)
        finally:
            if os.path.exists(fname):
                os.unlink(fname)


# ── Malformed tasks ───────────────────────────────────────────────────────────


class TestMalformedTasks:
    """Malformed entries at the task-dict level."""

    def test_task_missing_name_key(self):
        """Task without 'name' → empty string cell."""
        task = _valid_task()
        del task["name"]
        _generate([_valid_section(tasks=[task])])

    def test_task_name_is_none(self):
        """Task name=None → None stored in cell (no crash)."""
        _generate([_valid_section(tasks=[_valid_task(name=None)])])

    def test_task_time_estimate_is_string(self):
        """time_estimate='TBD' → stored in cell; not added to numeric total."""
        _generate([_valid_section(tasks=[_valid_task(time_estimate="TBD")])])

    def test_task_time_estimate_is_negative(self):
        """Negative time_estimate → accepted and totalled."""
        _generate([_valid_section(tasks=[_valid_task(time_estimate=-3.0)])])

    def test_task_time_estimate_is_zero(self):
        """Zero time_estimate → accepted."""
        _generate([_valid_section(tasks=[_valid_task(time_estimate=0)])])

    def test_task_progress_is_negative(self):
        """progress=-10 → DataBar clamps; no crash."""
        _generate([_valid_section(tasks=[_valid_task(progress=-10)])])

    def test_task_progress_over_100(self):
        """progress=150 → DataBar clamps; no crash."""
        _generate([_valid_section(tasks=[_valid_task(progress=150)])])

    def test_task_progress_is_string(self):
        """progress='N/A' → stored in cell; not added to numeric average."""
        _generate([_valid_section(tasks=[_valid_task(progress="N/A")])])

    def test_task_missing_ranges_key(self):
        """Task without 'ranges' key → no bars painted."""
        task = _valid_task()
        del task["ranges"]
        _generate([_valid_section(tasks=[task])])

    def test_task_ranges_is_none(self):
        """ranges=None → treated as empty list."""
        _generate([_valid_section(tasks=[_valid_task(ranges=None)])])

    def test_task_ranges_is_a_dict(self):
        """ranges is a single range dict (forgot to wrap) → treated as empty."""
        bad = {"start": START, "end": END, "color": "FF0000"}
        _generate([_valid_section(tasks=[_valid_task(ranges=bad)])])

    def test_task_missing_annotations_key(self):
        """Task without 'annotations' → no annotation cells filled."""
        task = _valid_task()
        del task["annotations"]
        _generate([_valid_section(tasks=[task])])


# ── Malformed ranges ──────────────────────────────────────────────────────────


class TestMalformedRanges:
    """Malformed entries inside a task's 'ranges' list."""

    def test_range_missing_start_key(self):
        """Range without 'start' → that range is skipped gracefully."""
        bad = {"end": END, "color": "FF0000"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[bad])])])

    def test_range_missing_end_key(self):
        """Range without 'end' → that range is skipped gracefully."""
        bad = {"start": START, "color": "FF0000"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[bad])])])

    def test_range_missing_both_date_keys(self):
        """Range with neither 'start' nor 'end' → skipped."""
        _generate([_valid_section(tasks=[_valid_task(ranges=[{"color": "FF0000"}])])])

    def test_range_is_none(self):
        """None entry inside ranges list → skipped."""
        good = {"start": START, "end": END, "color": "00B050"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[None, good])])])

    def test_range_is_a_string(self):
        """String entry inside ranges list → skipped."""
        good = {"start": START, "end": END}
        _generate([_valid_section(tasks=[_valid_task(ranges=["bad", good])])])

    def test_range_is_an_integer(self):
        """Integer entry inside ranges list → skipped."""
        _generate([_valid_section(tasks=[_valid_task(ranges=[42])])])

    def test_range_start_equals_end(self):
        """Single-day (milestone) range → exactly one bar cell."""
        day = datetime(2026, 4, 8)
        _generate([_valid_section(tasks=[_valid_task(ranges=[{"start": day, "end": day}])])])

    def test_range_reversed_dates(self):
        """start > end → empty iteration (range(x, y) where x>y), no crash."""
        bad = {"start": END, "end": START, "color": "FF0000"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[bad])])])

    def test_range_entirely_before_chart_start(self):
        """Range ends before chart start → no bar cells painted."""
        before = datetime(2026, 3, 1)
        _generate([_valid_section(tasks=[_valid_task(ranges=[{"start": before, "end": before}])])])

    def test_range_entirely_after_chart_end(self):
        """Range starts after chart end → no bar cells painted."""
        after = datetime(2026, 6, 1)
        _generate([_valid_section(tasks=[_valid_task(ranges=[{"start": after, "end": after}])])])

    def test_range_partially_outside_chart_left(self):
        """Range starts before chart start → clamped to chart start."""
        rng = {"start": datetime(2026, 3, 1), "end": END}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_partially_outside_chart_right(self):
        """Range ends after chart end → clamped to chart end."""
        rng = {"start": START, "end": datetime(2026, 7, 1)}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_spans_entire_chart_and_beyond(self):
        """Range fully wraps the chart window → all date cells painted."""
        rng = {"start": datetime(2026, 1, 1), "end": datetime(2026, 12, 31)}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_color_missing(self):
        """Range without 'color' → falls back to default bar_color."""
        _generate([_valid_section(tasks=[_valid_task(ranges=[{"start": START, "end": END}])])])

    def test_range_color_is_none(self):
        """Range color=None → falls back to default bar_color."""
        rng = {"start": START, "end": END, "color": None}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_color_invalid_hex_chars(self):
        """Non-hex color string → falls back to default bar_color."""
        rng = {"start": START, "end": END, "color": "ZZZZZZ"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_color_too_short(self):
        """3-char color string → falls back to default bar_color."""
        rng = {"start": START, "end": END, "color": "F00"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_range_color_with_hash_prefix(self):
        """Color with leading '#' (common user mistake) → handled or falls back."""
        rng = {"start": START, "end": END, "color": "#FF0000"}
        _generate([_valid_section(tasks=[_valid_task(ranges=[rng])])])

    def test_mixed_valid_and_invalid_ranges(self):
        """Mix of valid and invalid range entries in the same task → only valid ones painted."""
        ranges = [
            {"start": START, "end": END},               # valid – no color
            None,                                         # None entry
            {"color": "FF0000"},                          # missing dates
            {"start": END, "end": START},                 # reversed
            "string",                                     # non-dict
            {"start": START, "end": END, "color": "ABCDEF"},  # valid
        ]
        _generate([_valid_section(tasks=[_valid_task(ranges=ranges)])])


# ── Malformed annotations ─────────────────────────────────────────────────────


class TestMalformedAnnotations:
    """Malformed values for the 'annotations' key."""

    def test_annotations_is_none(self):
        """annotations=None → treated as empty dict."""
        _generate([_valid_section(tasks=[_valid_task(annotations=None)])])

    def test_annotations_is_a_list(self):
        """annotations is a list → treated as empty dict."""
        _generate([_valid_section(tasks=[_valid_task(annotations=["Alice", "Bob"])])])

    def test_annotations_is_a_string(self):
        """annotations is a string → treated as empty dict."""
        _generate([_valid_section(tasks=[_valid_task(annotations="Alice:A")])])

    def test_annotations_unknown_resource(self):
        """Resource key not in resource_names → entry silently ignored."""
        _generate([_valid_section(tasks=[_valid_task(annotations={"Unknown": "A"})])])

    def test_annotations_value_is_integer(self):
        """Annotation role value is an int → stored in cell, no crash."""
        _generate([_valid_section(tasks=[_valid_task(annotations={"Alice": 1})])])

    def test_annotations_value_is_none(self):
        """Annotation role value is None → stored as None, no crash."""
        _generate([_valid_section(tasks=[_valid_task(annotations={"Alice": None})])])

    def test_annotations_empty_dict(self):
        """Empty annotations → all annotation cells left blank."""
        _generate([_valid_section(tasks=[_valid_task(annotations={})])])


# ── Edge-case chart configuration ─────────────────────────────────────────────


class TestEdgeCaseChartConfig:
    """Edge cases at the GanttChart / date-range level."""

    def test_start_equals_end(self):
        """Single-day chart (start == end) → one date column."""
        d = datetime(2026, 4, 7)
        _generate([_valid_section()], start_date=d, end_date=d)

    def test_start_after_end(self):
        """start_date > end_date → empty date range; no crash."""
        _generate(
            [_valid_section()],
            start_date=datetime(2026, 4, 12),
            end_date=datetime(2026, 4, 6),
        )

    def test_project_name_is_none(self):
        """project_name=None → None stored in title cell, no crash."""
        _generate([_valid_section()], project_name=None)

    def test_resource_names_is_none(self):
        """resource_names=None → treated as empty list."""
        chart = GanttChart(
            sections=[_valid_section()],
            start_date=START,
            end_date=END,
            resource_names=None,
        )
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            fname = f.name
        try:
            chart.generate_excel(fname)
        finally:
            if os.path.exists(fname):
                os.unlink(fname)

    def test_all_tasks_have_no_progress(self):
        """No tasks supply progress → total progress cell left blank."""
        task = {
            "name": "No Progress Task",
            "time_estimate": 3,
            "ranges": [{"start": START, "end": END}],
        }
        _generate([{"name": "Sec", "tasks": [task]}])

    def test_all_tasks_have_no_time_estimate(self):
        """No numeric time_estimate → total = 0."""
        task = {
            "name": "No Est Task",
            "progress": 25,
            "ranges": [{"start": START, "end": END}],
        }
        _generate([{"name": "Sec", "tasks": [task]}])
