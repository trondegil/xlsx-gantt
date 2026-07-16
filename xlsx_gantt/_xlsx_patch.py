"""
xlsx_gantt/_xlsx_patch.py
~~~~~~~~~~~~~~~~~~~~~~~~~~
Post-processing utility that patches a saved ``.xlsx`` file so that its
DataBar conditional-format rules render as **solid** (no gradient) bars
using the Excel 2010 ``x14`` extension XML.

openpyxl writes ``<dataBar>`` elements without the ``x14`` gradient=0
attribute, which causes Excel to render faded/gradient bars instead of
the solid bars we want.  This module rewrites the raw ZIP entries for
the targeted worksheets to inject the required ``<x14:dataBar>``
extension nodes.

Two usage modes
---------------
File-based (default)::

    patch_solid_databars("gantt_chart.xlsx", "D4:D20")

In-memory (``BytesIO``)::

    buf = io.BytesIO(raw_xlsx_bytes)
    patch_solid_databars(buf, "D4:D20")
    result_bytes = buf.getvalue()

Both modes modify the data **in place** (file overwritten; ``BytesIO``
position reset to 0 so ``.getvalue()`` returns the patched content).
"""
from __future__ import annotations

import re
import uuid as _uuid
import zipfile
from io import BytesIO

_SHEET_ENTRY = re.compile(r"^xl/worksheets/[^/]+\.xml$")


def _patch_sheet_xml(xml: str, sqref: str) -> str:
    """Inject the x14 solid-DataBar extension into one worksheet's XML."""
    uid = str(_uuid.uuid4()).upper()

    # ── Inject x14 extension into the existing <dataBar> cfRule ──────
    id_ext = (
        "<extLst>"
        '<ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3B}" '
        'xmlns:x14="http://schemas.microsoft.com/office/'
        'spreadsheetml/2009/9/main">'
        f"<x14:id>{{{uid}}}</x14:id>"
        "</ext>"
        "</extLst>"
    )
    xml = re.sub(
        r"</dataBar>\s*</cfRule>",
        f"</dataBar>{id_ext}</cfRule>",
        xml,
        count=1,
    )

    # ── Append the top-level <x14:conditionalFormattings> extension ───
    x14 = (
        f'<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" '
        f'xmlns:x14="http://schemas.microsoft.com/office/'
        f'spreadsheetml/2009/9/main">'
        f"<x14:conditionalFormattings>"
        f"<x14:conditionalFormatting "
        f'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        f'<x14:cfRule type="dataBar" id="{{{uid}}}">'
        f'<x14:dataBar minLength="0" maxLength="100" gradient="0">'
        f'<x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo>'
        f'<x14:cfvo type="num"><xm:f>100</xm:f></x14:cfvo>'
        f"</x14:dataBar>"
        f"</x14:cfRule>"
        f"<xm:sqref>{sqref}</xm:sqref>"
        f"</x14:conditionalFormatting>"
        f"</x14:conditionalFormattings>"
        f"</ext>"
    )

    ext_lst_pos = xml.rfind("</extLst>")
    ws_end_pos  = xml.rfind("</worksheet>")
    if ext_lst_pos != -1 and ext_lst_pos > ws_end_pos - 30:
        xml = xml[:ext_lst_pos] + x14 + xml[ext_lst_pos:]
    else:
        xml = xml[:ws_end_pos] + f"<extLst>{x14}</extLst>" + xml[ws_end_pos:]

    return xml


def patch_solid_databars(
    target: str | BytesIO,
    sqref: str,
    *,
    sheet: str | None = "xl/worksheets/sheet1.xml",
) -> None:
    """
    Patch *target* so the DataBar CF rule on *sqref* renders as solid.

    Parameters
    ----------
    target:
        Either a file-system path (``str``) or an open :class:`io.BytesIO`
        object.  When a path is given the file is read, patched, and
        overwritten.  When a ``BytesIO`` is given it is patched in place
        and its position is reset to 0.
    sqref:
        The cell-range string that the DataBar rule was applied to
        (e.g. ``"D4:D40"``).
    sheet:
        Internal ZIP entry name for the worksheet that carries the
        DataBar rule.  Defaults to sheet 1.  Pass ``None`` to patch
        **every** worksheet that contains a DataBar rule.
    """
    # ── Read raw bytes ────────────────────────────────────────────────
    if isinstance(target, (str, bytes)):
        with open(target, "rb") as fh:
            raw = fh.read()
    else:
        target.seek(0)
        raw = target.read()

    # ── Unpack the ZIP ────────────────────────────────────────────────
    files: dict[str, bytes] = {}
    with zipfile.ZipFile(BytesIO(raw), "r") as zin:
        for name in zin.namelist():
            files[name] = zin.read(name)

    if sheet is None:
        targets = [
            name for name in files
            if _SHEET_ENTRY.match(name) and b"</dataBar>" in files[name]
        ]
    else:
        targets = [sheet] if sheet in files else []

    if not targets:
        return  # nothing to patch

    for name in targets:
        files[name] = _patch_sheet_xml(files[name].decode("utf-8"), sqref).encode("utf-8")

    # ── Repack the ZIP ────────────────────────────────────────────────
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    patched = out.getvalue()

    # ── Write back ───────────────────────────────────────────────────
    if isinstance(target, (str, bytes)):
        with open(target, "wb") as fh:
            fh.write(patched)
    else:
        target.seek(0)
        target.truncate()
        target.write(patched)
        target.seek(0)
