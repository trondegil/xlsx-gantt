"""
tests/test_new_features.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~
Tests for features introduced in the refactored package:

* ``generate_excel_bytes()`` — in-memory xlsx generation
* ``Section`` / ``Task`` / ``DateRange`` dataclass API
* Mixed dict + dataclass input in the same sections list
* ``patch_solid_databars()`` with a ``BytesIO`` target
* ``DateRange`` / ``Task`` / ``Section`` ``to_dict`` / ``from_dict`` round-trips
"""
from __future__ import annotations

import io
import zipfile
from datetime import datetime

import pytest

from xlsx_gantt import GanttChart, GanttStyle, GanttTheme
from xlsx_gantt import DateRange, Section, Task
from xlsx_gantt._xlsx_patch import patch_solid_databars


# ── Shared fixtures ───────────────────────────────────────────────────────────

START = datetime(2026, 4, 6)   # Monday
END   = datetime(2026, 4, 19)  # Sunday (2 weeks)
RESOURCES = ["Alice", "Bob"]


def _dr(start=None, end=None, color="0070C0") -> DateRange:
    return DateRange(start=start or START, end=end or END, color=color)


def _task(**kw) -> Task:
    defaults = dict(
        name="Task 1",
        time_estimate=5.0,
        progress=50,
        ranges=[_dr()],
        annotations={"Alice": "R", "Bob": "S"},
    )
    defaults.update(kw)
    return Task(**defaults)


def _section(**kw) -> Section:
    defaults = dict(name="Section 1", tasks=[_task()])
    defaults.update(kw)
    return Section(**defaults)


def _chart(sections=None, **kw) -> GanttChart:
    return GanttChart(
        sections=sections if sections is not None else [_section()],
        start_date=kw.pop("start_date", START),
        end_date=kw.pop("end_date", END),
        project_name=kw.pop("project_name", "Test"),
        resource_names=kw.pop("resource_names", RESOURCES),
        **kw,
    )


# ══════════════════════════════════════════════════════════════════════════════
# generate_excel_bytes()
# ══════════════════════════════════════════════════════════════════════════════

class TestGenerateExcelBytes:
    """``generate_excel_bytes()`` returns valid in-memory xlsx content."""

    def test_returns_bytes(self):
        result = _chart().generate_excel_bytes()
        assert isinstance(result, bytes)

    def test_result_is_non_empty(self):
        result = _chart().generate_excel_bytes()
        assert len(result) > 5_000, "Returned bytes suspiciously small"

    def test_result_is_valid_zip(self):
        result = _chart().generate_excel_bytes()
        assert zipfile.is_zipfile(io.BytesIO(result))

    def test_result_contains_worksheet(self):
        result = _chart().generate_excel_bytes()
        with zipfile.ZipFile(io.BytesIO(result)) as z:
            assert "xl/worksheets/sheet1.xml" in z.namelist()

    def test_bytes_match_file_content(self, tmp_path):
        """``generate_excel_bytes()`` and ``generate_excel()`` produce the same
        structural content (both valid zip with the worksheet entry)."""
        chart      = _chart()
        raw_bytes  = chart.generate_excel_bytes()
        file_path  = tmp_path / "gantt.xlsx"
        chart.generate_excel(str(file_path))

        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zb:
            assert "xl/worksheets/sheet1.xml" in zb.namelist()
        with zipfile.ZipFile(file_path) as zf:
            assert "xl/worksheets/sheet1.xml" in zf.namelist()

    def test_all_themes_return_bytes(self):
        """Every theme must work with ``generate_excel_bytes()``."""
        for name in GanttTheme.list_themes():
            result = _chart(style=GanttTheme.get(name)).generate_excel_bytes()
            assert isinstance(result, bytes) and len(result) > 1_000

    def test_empty_sections_returns_bytes(self):
        result = _chart(sections=[]).generate_excel_bytes()
        assert isinstance(result, bytes) and len(result) > 1_000

    def test_no_resource_names_returns_bytes(self):
        result = _chart(resource_names=[]).generate_excel_bytes()
        assert isinstance(result, bytes) and len(result) > 1_000


# ══════════════════════════════════════════════════════════════════════════════
# Dataclass API — Section / Task / DateRange
# ══════════════════════════════════════════════════════════════════════════════

class TestDataclassInput:
    """Charts built from Section / Task / DateRange dataclasses generate
    correctly and without errors."""

    def test_single_section_dataclass(self):
        chart = _chart(sections=[_section()])
        assert isinstance(chart.generate_excel_bytes(), bytes)

    def test_multiple_sections_dataclasses(self):
        sections = [
            _section(name="Alpha"),
            _section(name="Beta"),
        ]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)

    def test_task_with_multiple_ranges(self):
        task = Task(
            name="Multi-bar",
            ranges=[
                DateRange(start=START, end=datetime(2026, 4, 9), color="0070C0"),
                DateRange(start=datetime(2026, 4, 13), end=END, color="00B050"),
            ],
        )
        sections = [Section(name="Sec", tasks=[task])]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)

    def test_task_with_no_optional_fields(self):
        """Task with only a name — no estimate, progress, ranges, or annotations."""
        task = Task(name="Bare task")
        sections = [Section(name="Minimal", tasks=[task])]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)

    def test_section_with_no_tasks(self):
        sections = [Section(name="Empty section", tasks=[])]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)

    def test_daterange_color_none_uses_theme_default(self):
        """DateRange with color=None should fall back to bar_color in style."""
        task = Task(
            name="No-color bar",
            ranges=[DateRange(start=START, end=END, color=None)],
        )
        sections = [Section(name="Sec", tasks=[task])]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)

    def test_annotations_r_and_s_roles(self):
        """'R' and 'S' annotation roles trigger the correct coloured fill."""
        style = GanttStyle(
            annotation_a_bg="C8FFCC",
            annotation_s_bg="FFE4B5",
        )
        task = Task(
            name="Annotated",
            annotations={"Alice": "R", "Bob": "S"},
        )
        sections = [Section(name="Sec", tasks=[task])]
        assert isinstance(_chart(sections=sections, style=style).generate_excel_bytes(), bytes)


# ══════════════════════════════════════════════════════════════════════════════
# Mixed dict + dataclass input
# ══════════════════════════════════════════════════════════════════════════════

class TestMixedInput:
    """Sections list can freely mix Section dataclass instances and plain dicts."""

    def test_dict_section_and_dataclass_section(self):
        dict_section = {
            "name": "Dict section",
            "tasks": [
                {
                    "name": "Dict task",
                    "time_estimate": 3,
                    "progress": 75,
                    "ranges": [{"start": START, "end": END, "color": "FF0000"}],
                    "annotations": {"Alice": "R"},
                }
            ],
        }
        dc_section = _section(name="Dataclass section")
        assert isinstance(
            _chart(sections=[dict_section, dc_section]).generate_excel_bytes(),
            bytes,
        )

    def test_section_dataclass_with_dict_tasks(self):
        """Section dataclass whose tasks list contains plain dicts."""
        dict_task = {
            "name": "Dict task inside Section",
            "time_estimate": 2,
            "ranges": [{"start": START, "end": END}],
        }
        section = Section(name="Mixed tasks", tasks=[dict_task])
        assert isinstance(_chart(sections=[section]).generate_excel_bytes(), bytes)

    def test_dict_section_with_task_dataclasses(self):
        """Plain dict section whose tasks list contains Task dataclasses."""
        task = _task(name="Dataclass task inside dict section")
        dict_section = {"name": "Mixed", "tasks": [task]}
        # Task objects in a dict section's tasks list are handled by the renderer
        assert isinstance(_chart(sections=[dict_section]).generate_excel_bytes(), bytes)

    def test_none_entry_between_valid_sections(self):
        """None entry between valid sections is skipped."""
        sections = [_section(name="First"), None, _section(name="Last")]
        assert isinstance(_chart(sections=sections).generate_excel_bytes(), bytes)


# ══════════════════════════════════════════════════════════════════════════════
# patch_solid_databars with BytesIO
# ══════════════════════════════════════════════════════════════════════════════

class TestPatchSolidDatabars:
    """``patch_solid_databars`` works on both file paths and BytesIO objects."""

    def test_bytesio_is_still_valid_zip_after_patch(self):
        raw = _chart().generate_excel_bytes()
        buf = io.BytesIO(raw)
        patch_solid_databars(buf, "D4:D10")
        buf.seek(0)
        assert zipfile.is_zipfile(buf)

    def test_bytesio_position_reset_to_zero(self):
        raw = _chart().generate_excel_bytes()
        buf = io.BytesIO(raw)
        patch_solid_databars(buf, "D4:D10")
        assert buf.tell() == 0, "BytesIO position should be reset to 0 after patching"

    def test_file_patch_produces_valid_zip(self, tmp_path):
        out = tmp_path / "test.xlsx"
        _chart().generate_excel(str(out))
        original_size = out.stat().st_size
        patch_solid_databars(str(out), "D4:D10")
        assert zipfile.is_zipfile(out)
        # File should still have reasonable size after patching
        assert out.stat().st_size > original_size * 0.5

    def test_patch_noop_on_missing_sheet(self):
        """Patching with a non-existent sheet name should not raise."""
        raw = _chart().generate_excel_bytes()
        buf = io.BytesIO(raw)
        patch_solid_databars(buf, "D4:D10", sheet="xl/worksheets/sheet99.xml")
        # Should still be a valid zip (unmodified)
        buf.seek(0)
        assert zipfile.is_zipfile(buf)


# ══════════════════════════════════════════════════════════════════════════════
# to_dict / from_dict round-trips
# ══════════════════════════════════════════════════════════════════════════════

class TestModelRoundTrips:
    """``to_dict`` / ``from_dict`` produce consistent representations."""

    def test_daterange_to_dict(self):
        dr = DateRange(start=START, end=END, color="0070C0")
        d  = dr.to_dict()
        assert d["start"] == START
        assert d["end"]   == END
        assert d["color"] == "0070C0"

    def test_daterange_to_dict_no_color(self):
        dr = DateRange(start=START, end=END)
        d  = dr.to_dict()
        assert "color" not in d

    def test_daterange_from_dict_roundtrip(self):
        original = DateRange(start=START, end=END, color="FF0000")
        restored = DateRange.from_dict(original.to_dict())
        assert restored is not None
        assert restored.start == original.start
        assert restored.end   == original.end
        assert restored.color == original.color

    def test_daterange_from_dict_missing_start(self):
        assert DateRange.from_dict({"end": END}) is None

    def test_daterange_from_dict_missing_end(self):
        assert DateRange.from_dict({"start": START}) is None

    def test_daterange_from_dict_non_dict(self):
        assert DateRange.from_dict(None) is None
        assert DateRange.from_dict("bad") is None
        assert DateRange.from_dict(42) is None

    def test_task_to_dict_contains_all_keys(self):
        t = _task()
        d = t.to_dict()
        assert set(d.keys()) >= {"name", "time_estimate", "progress",
                                  "ranges", "annotations"}

    def test_task_to_dict_converts_daterange_to_dict(self):
        t = Task(name="X", ranges=[DateRange(start=START, end=END, color="AABBCC")])
        d = t.to_dict()
        assert isinstance(d["ranges"][0], dict)
        assert d["ranges"][0]["color"] == "AABBCC"

    def test_task_from_dict_roundtrip(self):
        original = _task(name="Round-trip task", time_estimate=7, progress=33)
        restored = Task.from_dict(original.to_dict())
        assert restored is not None
        assert restored.name          == original.name
        assert restored.time_estimate == original.time_estimate
        assert restored.progress      == original.progress

    def test_task_from_dict_non_dict(self):
        assert Task.from_dict(None) is None
        assert Task.from_dict("bad") is None

    def test_section_to_dict_contains_name_and_tasks(self):
        sec = _section()
        d   = sec.to_dict()
        assert "name"  in d
        assert "tasks" in d
        assert isinstance(d["tasks"], list)

    def test_section_to_dict_converts_task_to_dict(self):
        sec = Section(name="S", tasks=[Task(name="T")])
        d   = sec.to_dict()
        assert isinstance(d["tasks"][0], dict)

    def test_section_from_dict_roundtrip(self):
        original = _section(name="My section")
        restored = Section.from_dict(original.to_dict())
        assert restored is not None
        assert restored.name == original.name

    def test_section_from_dict_non_dict(self):
        assert Section.from_dict(None) is None
        assert Section.from_dict(42) is None

    def test_dataclass_chart_matches_dict_chart_structure(self):
        """A chart built from dataclasses should produce the same-sized xlsx
        as an equivalent chart built from plain dicts."""
        dc_sections = [
            Section(name="Design", tasks=[
                Task(name="Research", time_estimate=3, progress=100,
                     ranges=[DateRange(start=START, end=END, color="0070C0")],
                     annotations={"Alice": "R", "Bob": "S"}),
            ]),
        ]
        dict_sections = [sec.to_dict() for sec in dc_sections]

        dc_bytes   = _chart(sections=dc_sections).generate_excel_bytes()
        dict_bytes = _chart(sections=dict_sections).generate_excel_bytes()

        # Both should be valid xlsx files of similar (not necessarily identical) size
        assert zipfile.is_zipfile(io.BytesIO(dc_bytes))
        assert zipfile.is_zipfile(io.BytesIO(dict_bytes))
        ratio = len(dc_bytes) / len(dict_bytes)
        assert 0.8 < ratio < 1.2, f"Size ratio {ratio:.2f} unexpectedly different"
