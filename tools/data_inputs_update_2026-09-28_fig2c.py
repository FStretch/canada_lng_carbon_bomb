"""Data Inputs workbook update, 28 Sep 2026.

Figure 2(c) reads canada_2030_projected (646 Mt, current-policy 2030
projection). The row was reference-only. This marks it live and says so in
the note. Source and URL are left as they stand.

Asserts the row's current state before writing.

Usage, from the repository root:

    python tools/data_inputs_update_2026-09-28_fig2c.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DEFAULT = ROOT / "Inputs" / "Canada_LNG_Data_Inputs.xlsx"

NOTE = (
    "LIVE. Read by src/figures_report.py figure_2c_territorial_trajectory as "
    "the projected 2030 emissions under current policy, the single marker on "
    "Figure 2(c). The government's own projection of where 2030 emissions "
    "actually land, against a target of 417-455, and the evidence that Canada "
    "is not on track. canada_2030_overshoot_gap = 200 remains the overshoot line."
)
OLD_NOTE = (
    "REFERENCE-ONLY. Read by no module. The government's own projection of "
    "where 2030 emissions actually land, against a target of 417-455. Kept "
    "because it is the evidence for the statement that Canada is not on "
    "track. LIVE COUNTERPART: canada_2030_overshoot_gap = 200, which is the "
    "figure the model reads for the overshoot line."
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
    assert pa.cell(r, h["status"]).value == "reference-only", pa.cell(r, h["status"]).value
    assert pa.cell(r, h["notes"]).value == OLD_NOTE, pa.cell(r, h["notes"]).value
    assert pa.cell(r, h["source_url"]).value == SOURCE_URL
    source = pa.cell(r, h["source"]).value
    assert source and "current policy" in str(source)
    pa.cell(r, h["status"]).value = "live"
    pa.cell(r, h["notes"]).value = NOTE
    assert pa.cell(r, h["source"]).value == source
    assert pa.cell(r, h["source_url"]).value == SOURCE_URL
    assert pa.cell(r, h["value"]).value == 646
    wb.save(path)
    print("Parameters canada_2030_projected: reference-only -> live; note names figure 2(c)")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(DEFAULT))
