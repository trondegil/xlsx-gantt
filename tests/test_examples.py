"""
Ensures every plan in ``examples/`` loads with ``GanttChart.from_json`` and
renders, so the files linked from README.md stay valid.
"""
from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from xlsx_gantt import GanttChart

EXAMPLES = sorted((Path(__file__).parent.parent / "examples").glob("*.json"))


def test_examples_exist():
    assert EXAMPLES


@pytest.mark.parametrize("path", EXAMPLES, ids=lambda p: p.name)
def test_example_renders(path):
    chart = GanttChart.from_json(str(path))
    assert zipfile.is_zipfile(BytesIO(chart.generate_excel_bytes()))
