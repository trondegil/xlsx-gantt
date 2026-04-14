"""
xlsx_gantt/models.py
~~~~~~~~~~~~~~~~~~~~~
Typed input dataclasses for :class:`~xlsx_gantt.GanttChart`.

Using these dataclasses is **optional** — :class:`GanttChart` also
accepts plain :class:`dict` objects (the original API) and will
auto-coerce them via :meth:`Section.from_dict` internally.  The
dataclasses provide a type-safe, IDE-friendly alternative::

    from datetime import datetime
    from xlsx_gantt import DateRange, Task, Section, GanttChart

    task = Task(
        name          = "Backend",
        time_estimate = 10,
        progress      = 50,
        ranges        = [
            DateRange(
                start = datetime(2026, 1, 19),
                end   = datetime(2026, 2, 6),
                color = "00B050",
            )
        ],
        annotations   = {"Alice": "S", "Bob": "R"},
    )

    section = Section(name="Development", tasks=[task])

    chart = GanttChart(sections=[section], ...)

You can freely mix dataclass instances and raw dicts in the same
``sections`` list — both are handled transparently.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


# ══════════════════════════════════════════════════════════════════════
# DateRange
# ══════════════════════════════════════════════════════════════════════

@dataclass
class DateRange:
    """
    A single coloured bar on the Gantt timeline.

    Parameters
    ----------
    start:
        Inclusive start date.
    end:
        Inclusive end date.  ``start == end`` renders a one-cell
        *milestone* bar.
    color:
        6-digit hex string (no ``#``), e.g. ``"0070C0"``.
        Falls back to the chart's default ``bar_color`` if omitted or
        ``None``.
    """

    start: datetime
    end:   datetime
    color: str | None = None

    def to_dict(self) -> dict:
        """Return a plain-dict representation compatible with the dict API."""
        d: dict = {"start": self.start, "end": self.end}
        if self.color is not None:
            d["color"] = self.color
        return d

    @classmethod
    def from_dict(cls, data: object) -> "DateRange | None":
        """
        Construct from a raw dict.  Returns ``None`` for any non-dict
        input (the chart renderer will skip it gracefully).
        """
        if not isinstance(data, dict):
            return None
        if "start" not in data or "end" not in data:
            return None
        return cls(
            start=data["start"],
            end=data["end"],
            color=data.get("color"),
        )


# ══════════════════════════════════════════════════════════════════════
# Task
# ══════════════════════════════════════════════════════════════════════

@dataclass
class Task:
    """
    A single row in the Gantt chart.

    Parameters
    ----------
    name:
        Label shown in the *Task* column.
    time_estimate:
        Numeric value in any unit (hours, days, …).  Summed in the
        *Total* footer row.  Pass ``None`` to leave the cell blank.
    progress:
        Completion percentage (0–100).  Rendered as a solid DataBar.
        Pass ``None`` to leave the cell blank.
    ranges:
        One or more :class:`DateRange` objects (or plain dicts) that
        define the coloured bars on the timeline.
    annotations:
        Mapping of resource-name → role string (``"R"`` / ``"S"`` or
        any short label).  Keys not present in
        :attr:`GanttChart.resource_names` are silently ignored.
    """

    name:          str
    time_estimate: float | None                 = None
    progress:      float | None                 = None
    ranges:        list[DateRange | dict]       = field(default_factory=list)
    annotations:   dict[str, str]               = field(default_factory=dict)

    def to_dict(self) -> dict:
        """Return a plain-dict representation compatible with the dict API."""
        return {
            "name":          self.name,
            "time_estimate": self.time_estimate,
            "progress":      self.progress,
            "ranges": [
                r.to_dict() if isinstance(r, DateRange) else r
                for r in self.ranges
            ],
            "annotations": self.annotations,
        }

    @classmethod
    def from_dict(cls, data: object) -> "Task | None":
        """
        Construct from a raw dict.  Returns ``None`` for any non-dict
        input so the chart renderer can skip it gracefully.
        """
        if not isinstance(data, dict):
            return None
        raw_ranges = data.get("ranges") or []
        if not isinstance(raw_ranges, list):
            raw_ranges = []
        ranges = [r for r in raw_ranges]   # keep as-is; renderer handles guards
        return cls(
            name          = data.get("name", ""),
            time_estimate = data.get("time_estimate"),
            progress      = data.get("progress"),
            ranges        = ranges,
            annotations   = data.get("annotations") or {},
        )


# ══════════════════════════════════════════════════════════════════════
# Section
# ══════════════════════════════════════════════════════════════════════

@dataclass
class Section:
    """
    A group of tasks shown under a common *Activity* label.

    The section name is merged vertically across all task rows in
    column A.

    Parameters
    ----------
    name:
        Activity label (column A).
    tasks:
        Ordered list of :class:`Task` objects (or plain dicts).
    """

    name:  str
    tasks: list[Task | dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        """Return a plain-dict representation compatible with the dict API."""
        return {
            "name": self.name,
            "tasks": [
                t.to_dict() if isinstance(t, Task) else t
                for t in self.tasks
            ],
        }

    @classmethod
    def from_dict(cls, data: object) -> "Section | None":
        """
        Construct from a raw dict.  Returns ``None`` for non-dict input.
        """
        if not isinstance(data, dict):
            return None
        raw_tasks = data.get("tasks", [])
        if not isinstance(raw_tasks, list):
            raw_tasks = []
        return cls(
            name  = data.get("name", ""),
            tasks = raw_tasks,   # kept as-is for renderer guards
        )
