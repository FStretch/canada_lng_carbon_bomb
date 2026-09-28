"""Data Inputs workbook update, 28 Sep 2026.

Figure 2(c) no longer reads canada_2030_projected. Return the row to
reference-only and restore the earlier note, recording the removal.

Asserts the row is currently live before writing.

Usage, from the repository root:

    python tools/data_inputs_update_2026-09-28_fig2c_revert.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DEFAULT = ROOT / "Inputs" / "Canada_LNG_Data_Inputs.xlsx"

LIVE_NOTE = (
    "LIVE. Read by src/figures_report.py figure_2c_territorial_trajectory as "
    "the projected 2030 emissions under current policy, the single marker on "
    "Figure 2(c). The government's own projection of where 2030 emissions "
    "actually land, against a target of 417-455, and the evidence that Canada "
    "is not on track. canada_2030_overshoot_gap = 200 remains the overshoot line."
)
NOTE = (
    "REFERENCE-ONLY. Read by no module. The government's own projection of "
    "where 2030 emissions actually land, against a target of 417-455. Kept "
    "because it is the evidence for the statement that Canada is not on "
    "track. LIVE COUNTERPART: canada_2030_overshoot_gap = 200, which is the "
    "figure the model reads for the overshoot line. "
    "Read by no module (removed from figure 2(c) on 28 Sep 2026)."
)
SOURCE_URL = "https://climateactiontracker.org/countries/canada/"


def find_row(ws, key: str, col: int = 1) -> int:
    hits = [r for r in range(1, ws.max_row + 1) if ws.cell(r, col).value == key]
    assert len(hits) == 1, (ws.title, key, hits)
    return hits[0]


def hdr(ws) -> dict:
    return {c.value: i for i, c in enumerate(ws[1], start=1) if c.value}


def main(path: str) -> None:
    wb = openpyxl.load_workbook(path)
    pa = wb["Parameters"]
    h = hdr(pa)
    r = find_row(pa, "canada_2030_projected")
    assert pa.cell(r, h["value"]).value == 646, pa.cell(r, h["value"]).value
    assert pa.cell(r, h["status"]).value == "live", pa.cell(r, h["status"]).value
    assert pa.cell(r, h["notes"]).value == LIVE_NOTE, pa.cell(r, h["notes"]).value
    assert pa.cell(r, h["source_url"]).value == SOURCE_URL
    source = pa.cell(r, h["source"]).value
    assert source and "current policy" in str(source)
    pa.cell(r, h["status"]).value = "reference-only"
    pa.cell(r, h["notes"]).value = NOTE
    assert pa.cell(r, h["source"]).value == source
    assert pa.cell(r, h["source_url"]).value == SOURCE_URL
    assert pa.cell(r, h["value"]).value == 646
    assert pa.cell(r, h["status"]).value == "reference-only"
    wb.save(path)
    print(
        "Parameters canada_2030_projected: live -> reference-only; "
        "note records removal from figure 2(c)"
    )


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(DEFAULT))
