"""
xlsx_gantt/themes.py
~~~~~~~~~~~~~~~~~~~~~
Color-theory helpers and pre-built GanttStyle themes.

All color values are 6-digit hex strings without a leading ``#``
(e.g. ``"FFC000"``).  The ``from_color`` factory accepts an optional
leading ``#`` and normalises it internally.
"""
from __future__ import annotations

import colorsys
from collections.abc import Callable

from .style import GanttStyle

# ══════════════════════════════════════════════════════════════════════
# Low-level color helpers
# ══════════════════════════════════════════════════════════════════════

def _parse(hex6: str) -> tuple[float, float, float]:
    """Hex string → (r, g, b) each in 0–1 range."""
    h = hex6.lstrip("#")
    return int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0


def _fmt(r: float, g: float, b: float) -> str:
    """(r, g, b) each 0–1 → upper-case 6-digit hex string."""
    return f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"


def _luminance(hex6: str) -> float:
    """Relative luminance as defined by WCAG 2.1 (0 = black, 1 = white)."""
    def _lin(c: float) -> float:
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _parse(hex6)
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def contrast_text(hex6: str) -> str:
    """
    Return ``"FFFFFF"`` (white) or ``"000000"`` (black) — whichever achieves
    better WCAG contrast ratio against *hex6*.

    The cross-over luminance threshold is ``≈ 0.179``.
    """
    return "FFFFFF" if _luminance(hex6) < 0.179 else "000000"


def darken(hex6: str, amount: float) -> str:
    """
    Blend *hex6* toward pure black.

    *amount* is in the range ``[0, 1]``; ``0`` → unchanged, ``1`` → ``"000000"``.
    """
    r, g, b = _parse(hex6.lstrip("#"))
    t = max(0.0, 1.0 - amount)
    return _fmt(r * t, g * t, b * t)


def lighten(hex6: str, amount: float) -> str:
    """
    Blend *hex6* toward pure white.

    *amount* is in the range ``[0, 1]``; ``0`` → unchanged, ``1`` → ``"FFFFFF"``.
    """
    r, g, b = _parse(hex6.lstrip("#"))
    a = max(0.0, min(1.0, amount))
    return _fmt(r + (1.0 - r) * a, g + (1.0 - g) * a, b + (1.0 - b) * a)


def rotate_hue(hex6: str, degrees: float) -> str:
    """
    Rotate the hue of *hex6* by *degrees* (0–360), preserving
    its original lightness and saturation.
    """
    r, g, b = _parse(hex6.lstrip("#"))
    h, lit, s = colorsys.rgb_to_hls(r, g, b)
    h = (h + degrees / 360.0) % 1.0
    return _fmt(*colorsys.hls_to_rgb(h, lit, s))


def _vivid_mid(hex6: str) -> str:
    """
    Return a vivid, mid-lightness color at the **same hue** as *hex6*.

    Lightness is pinned to 0.45 and saturation to 0.65, producing a
    solid, readable bar colour regardless of whether the source is very
    dark or very light.
    """
    r, g, b = _parse(hex6.lstrip("#"))
    h, _l, _s = colorsys.rgb_to_hls(r, g, b)
    return _fmt(*colorsys.hls_to_rgb(h, 0.45, 0.65))


# ══════════════════════════════════════════════════════════════════════
# Theme factory
# ══════════════════════════════════════════════════════════════════════

class GanttTheme:
    """
    Creates :class:`~xlsx_gantt.GanttStyle` instances from a base colour
    or by selecting a named pre-built theme.

    Quick-start::

        from xlsx_gantt import GanttTheme

        # Pre-built theme by method
        style = GanttTheme.ocean()

        # Pre-built theme by name
        style = GanttTheme.get("ocean")

        # Custom theme from any hex colour
        style = GanttTheme.from_color("1A5276")

        # See all theme names
        print(GanttTheme.list_themes())

    The ``bar_color`` field in the generated style is the **default** bar
    fill.  Any ``'color'`` key set on individual task ranges in your event
    data will override it (existing behaviour is unchanged).
    """

    # ── Algorithmic constructor ────────────────────────────────────────

    @staticmethod
    def from_color(base: str) -> GanttStyle:
        """
        Derive a complete :class:`GanttStyle` from a single *base* hex color.

        Derivation rules
        ----------------
        * **Header rows** (logo, week bands, day names): background = *base*;
          text auto-chosen for WCAG contrast (white on dark, black on light).
        * **Column-label & Total rows**: background = *base* darkened 88 %
          toward black; text auto-chosen (almost always white).
        * **Info columns** (Time Est. / Progress): very pale tint
          (88 % toward white).
        * **Default bar fill**: vivid mid-tone at the same hue as *base*.
        * **Progress data-bar**: *base* darkened 25 % — keeps it readable.
        * Weekend & border colours are neutral (``EFEFEF`` / ``CCCCCC``).

        Parameters
        ----------
        base:
            6-digit hex color string, with or without a leading ``#``.
        """
        base = base.lstrip("#").upper()

        hdr_fg   = contrast_text(base)        # white or black on header
        dark_bg  = darken(base, 0.88)         # near-black tinted shade
        dark_fg  = contrast_text(dark_bg)     # almost always white
        even_bg  = lighten(base, 0.88)        # very pale info-column tint
        bar      = _vivid_mid(base)           # vivid same-hue bar default
        prog_bar = lighten(base, 0.45)        # lighter progress bar tint

        hdr_border = darken(base, 0.20)   # slightly darker than the header bg

        return GanttStyle(
            # ── Primary palette ──────────────────────────────────────
            header_bg           = base,
            header_fg           = hdr_fg,
            weekend_bg          = "EFEFEF",
            bar_color           = bar,
            progress_bar_color  = prog_bar,
            border_color        = "CCCCCC",
            header_border_color = hdr_border,
            # ── Header sub-colours ────────────────────────────────────
            logo_bg            = base,
            title_fg           = hdr_fg,
            week_label_bg      = base,
            week_label_fg      = hdr_fg,
            day_name_bg        = base,
            day_name_fg        = hdr_fg,
            date_number_fg     = dark_fg,
            col_label_bg       = dark_bg,
            col_label_fg       = dark_fg,
            # ── Data row colours ──────────────────────────────────────
            row_bg_even        = None,
            section_name_bg    = None,
            col_info_bg        = even_bg,
            section_name_fg    = "000000",
            task_fg            = "000000",
            estimate_fg        = "000000",
            progress_value_fg  = "000000",
            # ── Total row ─────────────────────────────────────────────
            total_row_bg       = dark_bg,
            total_row_fg       = dark_fg,
        )

    # ── Pre-built named themes ─────────────────────────────────────────

    @staticmethod
    def amber() -> GanttStyle:
        """Warm amber / corporate orange.  (base ``FFC000``)"""
        s = GanttTheme.from_color("FFC000")
        # Force white header text — the WCAG algorithm picks black on bright amber,
        # but white reads better against the saturated background.
        s.header_fg     = "FFFFFF"
        s.title_fg      = "FFFFFF"
        s.week_label_fg = "FFFFFF"
        s.day_name_fg   = "FFFFFF"
        return s

    @staticmethod
    def ocean() -> GanttStyle:
        """Deep ocean blue — professional and calm.  (base ``1A5276``)"""
        return GanttTheme.from_color("1A5276")

    @staticmethod
    def forest() -> GanttStyle:
        """Forest green — environmental / health-sector projects.  (base ``1E8449``)"""
        return GanttTheme.from_color("1E8449")

    @staticmethod
    def crimson() -> GanttStyle:
        """Bold crimson — urgent, high-impact projects.  (base ``922B21``)"""
        return GanttTheme.from_color("922B21")

    @staticmethod
    def slate() -> GanttStyle:
        """Cool slate gray — minimal, neutral look.  (base ``5D6D7E``)"""
        return GanttTheme.from_color("5D6D7E")

    @staticmethod
    def royal_purple() -> GanttStyle:
        """Royal purple — the original default palette.  (base ``7030A0``)"""
        return GanttTheme.from_color("7030A0")

    @staticmethod
    def midnight() -> GanttStyle:
        """Near-black midnight — modern, dramatic header.  (base ``1B2631``)"""
        return GanttTheme.from_color("1B2631")

    @staticmethod
    def teal() -> GanttStyle:
        """Teal / blue-green — fresh and contemporary.  (base ``148F77``)"""
        s = GanttTheme.from_color("148F77")
        # Force white header text — the WCAG algorithm picks black on medium teal,
        # but white reads better against the saturated background.
        s.header_fg     = "FFFFFF"
        s.title_fg      = "FFFFFF"
        s.week_label_fg = "FFFFFF"
        s.day_name_fg   = "FFFFFF"
        return s

    # ── Registry & lookup ─────────────────────────────────────────────

    #: Mapping of lower-case theme names → factory callables.
    #: Populated below after the class body is complete.
    _REGISTRY: dict[str, Callable[[], GanttStyle]] = {}

    @classmethod
    def get(cls, name: str) -> GanttStyle:
        """
        Return a pre-built theme by *name* (case-insensitive; spaces and
        hyphens are treated the same as underscores).

        Raises ``ValueError`` for an unrecognised name.
        """
        key = name.lower().replace(" ", "_").replace("-", "_")
        fn  = cls._REGISTRY.get(key)
        if fn is None:
            raise ValueError(
                f"Unknown theme {name!r}.  "
                f"Available: {', '.join(sorted(cls._REGISTRY))}"
            )
        return fn()

    @classmethod
    def list_themes(cls) -> list[str]:
        """Return the names of all pre-built themes, sorted alphabetically."""
        return sorted(cls._REGISTRY)


# Populate registry after the class body is fully defined
GanttTheme._REGISTRY = {
    "amber":        GanttTheme.amber,
    "ocean":        GanttTheme.ocean,
    "forest":       GanttTheme.forest,
    "crimson":      GanttTheme.crimson,
    "slate":        GanttTheme.slate,
    "royal_purple": GanttTheme.royal_purple,
    "midnight":     GanttTheme.midnight,
    "teal":         GanttTheme.teal,
}
