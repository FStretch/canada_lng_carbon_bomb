"""Data Inputs workbook update, 27 Sep 2026.

Decisions (Freddie, 27 Sep 2026):
  2. Sample liquefaction in the Monte Carlo on a gas-turbine triangle 0.26 / 0.29 / 0.36.
  4. Electric drive is a separate sensitivity with every terminal electric (all gas vs all electric).
  1. The FID delay moves a project_life_years life later instead of cutting it.
Plus tidy items: stale 0.12 text, the 50-year licence note.

Usage: python data_inputs_update_2026-09-27.py path/to/Canada_LNG_Data_Inputs.xlsx
"""

from __future__ import annotations

import sys

import openpyxl

DELPHI = ("https://www2.gov.bc.ca/assets/gov/environment/climate-change/ind/lng/"
          "lng_emissions_benchmarking_-_march_2013.pdf")
PEMBINA = "https://www.pembina.org/reports/squaring-the-circle-state-of-lng-2023.pdf"
IEA = ("https://iea.blob.core.windows.net/assets/df9b1bed-4e16-4db8-bec5-a5fa117a9130/"
       "AssessingemissionsfromLNGsupplyandabatementoptions.pdf")
EAO_KSI = "https://iaac-aeic.gc.ca/050/documents/p82797/163192E.pdf"
MPO_LNGC = "https://www.canada.ca/en/privy-council/major-projects-office/projects/national/lng-canada.html"


def find_row(ws, key: str, col: int = 1) -> int:
    hits = [r for r in range(1, ws.max_row + 1) if ws.cell(r, col).value == key]
    assert len(hits) == 1, (ws.title, key, hits)
    return hits[0]


def hdr(ws) -> dict:
    return {c.value: i for i, c in enumerate(ws[1], start=1) if c.value}


def main(path: str):
    wb = openpyxl.load_workbook(path)

    # --- Emission Factors: liquefaction ---------------------------------
    ef = wb["Emission Factors"]
    h = hdr(ef)
    r = find_row(ef, "liquefaction")
    assert ef.cell(r, h["central"]).value == 0.29
    assert ef.cell(r, h["range_low"]).value == 0.15
    assert ef.cell(r, h["range_high"]).value == 0.36
    ef.cell(r, h["range_low"]).value = 0.26
    ef.cell(r, h["what_happens"]).value = (
        "Gas is chilled to minus 162 degrees at the terminal until it becomes liquid. Energy "
        "intensive. The central factor is gas-turbine drive for every terminal. The Monte Carlo "
        "samples 0.26 to 0.36, the range for gas-turbine plants. Electric drive is not in this "
        "range: it is a separate sensitivity with every terminal electric (Parameters "
        "liquefaction_electric 0.15 and liquefaction_electric_drive_grid 0.021).")
    ef.cell(r, h["range_sources"]).value = (
        "Both ends are gas-turbine plants in Delphi Group (2013), LNG Emissions Benchmarking, "
        f"prepared for the BC Ministry of Environment, March 2013: {DELPHI} . "
        "range_low 0.26 tCO2e per t LNG: Sabine Pass, the lowest gas-turbine plant in the "
        "benchmark (page 23, Figure 3). Adopted 27 September 2026, replacing 0.15. "
        "range_high 0.36 tCO2e per t LNG: Pluto LNG, a gas-turbine plant (GE Frame 7EA), in a "
        "benchmark of operating and proposed plants that ranged from 0.17 to 0.49 (page 14). "
        "Cited 27 September 2026; previously an uncited literature upper bound. "
        "The Monte Carlo samples liquefaction on the triangle 0.26 / 0.29 / 0.36 from 27 "
        "September 2026 (decision: Freddie). Electric drive is deliberately outside this range. "
        "The former range_low 0.15 (Pembina Institute, Gorski and Lam 2023, Squaring the Circle: "
        f"State of LNG 2023, {PEMBINA} , LNG Canada under electric drive) now lives only on the "
        "Parameters sheet as liquefaction_electric and drives the all-electric sensitivity in "
        "src/drive_sensitivity.py. Corroboration on the central: IEA (2025), page 13, gives a "
        "global average liquefaction intensity of about 6 g CO2-eq/MJ of LNG, which at 55 MJ/kg "
        f"is 0.33 tCO2e/t - above this model's 0.29, so 0.29 is not a high-side choice. {IEA}")
    print("Emission Factors liquefaction: range_low 0.15 -> 0.26, sources cited")

    # --- Parameters --------------------------------------------------------
    pa = wb["Parameters"]
    h = hdr(pa)
    r = find_row(pa, "liquefaction_electric")
    assert pa.cell(r, h["value"]).value == 0.15
    pa.cell(r, h["notes"]).value = (
        "Electric-drive sensitivity (src/drive_sensitivity.py): every headline terminal on "
        "electric drive at this figure, against the all-gas central (decision 27 Sep 2026). Also "
        "the section 9 electrification appendix. No longer Emission Factors range_low, which is "
        "0.26 (gas-turbine plants only) from 27 Sep 2026. NOT used in the central case, which is "
        "gas turbine 0.29 for every terminal.")
    r = find_row(pa, "liquefaction_electric_drive_grid")
    assert pa.cell(r, h["value"]).value == 0.021
    pa.cell(r, h["notes"]).value = (
        "SENSITIVITY ONLY - the central liquefaction factor stays 0.29 for every terminal. Used "
        "as the best case in the all-electric drive sensitivity (src/drive_sensitivity.py, "
        "decision 27 Sep 2026). BOUNDARY: this is the facility total intensity including marine "
        "sources, not the liquefaction stage alone, so it is not like for like with the 0.29 "
        "liquefaction factor on the Emission Factors sheet. Any comparison against 0.29 is "
        "therefore approximate and is labelled as such wherever it is reported.")
    r = find_row(pa, "fid_delay_mid")
    pa.cell(r, h["notes"]).value = (
        "Unchanged at 5, so the central case is untouched by the 3 September 2026 band change. "
        "Applied to pre-FID headline assets. Two exceptions from 27 September 2026 (decisions: "
        "Freddie). (1) Not added where the register's first_export_year_assumes_fid is True, "
        "because the developer date already allows for the build after an FID: Ksi Lisims "
        "(2032) and Tilbury Phase 1b (2031). (2) Where an asset's life comes from "
        "project_life_years with no licence end and no authorised term (Fermeuse 18 years, "
        "Summit Lake 30 years), the delay moves the whole life later instead of cutting its "
        "last years; the Monte Carlo does the same draw by draw. post_fid_to_completion = 4.5 "
        "rests on the same IEA statement and is reference-only; this row is the one the model "
        "reads.")
    print("Parameters: liquefaction_electric, liquefaction_electric_drive_grid, fid_delay_mid notes")

    # --- Scenarios: high_50yr ---------------------------------------------
    sc = wb["Scenarios"]
    r = find_row(sc, "high_50yr")
    assert sc.cell(r, 5).value == "CER may issue LNG licences up to 50 years."
    sc.cell(r, 5).value = (
        "50 years is the legal maximum term for a CER LNG export licence since Bill C-15 "
        "received royal assent on 26 March 2026. Source: Major Projects Office, LNG Canada "
        f"Phase 2: {MPO_LNGC}")
    print("Scenarios high_50yr note updated")

    # --- README sheet ------------------------------------------------------
    rd = wb["README"]
    r = find_row(rd, "Liquefaction is a single central value")
    rd.cell(r, 1).value = "Liquefaction: one central value, gas-turbine range"
    rd.cell(r, 2).value = (
        "The liquefaction stage carries a central factor of 0.29 tCO2e per tonne (gas turbine) "
        "for every terminal and its whole operating life. The Monte Carlo samples 0.26 / 0.29 / "
        "0.36, the range for gas-turbine plants in Delphi Group (2013) (Sabine Pass 0.26, Pluto "
        "0.36). Electric drive is not assumed for any project; it is a separate sensitivity with "
        "every terminal electric at 0.15 (Pembina, Gorski and Lam 2023) and 0.021 (BC EAO, Ksi "
        "Lisims grid case). The two together show all gas against all electric.")
    r = find_row(rd, "One central value per stage")
    rd.cell(r, 2).value = (
        "Each stage carries a single central figure. Low and high are not used to build "
        "scenarios; they set the triangle each stage is sampled on in the Monte Carlo. The "
        "stages are not correlated: combustion varies by about 9 per cent with gas composition, "
        "while upstream varies by a factor of three because measurement is difficult. Moving "
        "all six together would produce a spread wider than anything physically plausible.")
    r = find_row(rd, "Build")
    rd.cell(r, 2).value = "2026-09-27 (inputs update; see Emission Factors liquefaction and Parameters fid_delay_mid)"
    print("README sheet rows updated")

    # --- Data Gaps -----------------------------------------------------------
    dg = wb["Data Gaps"]
    h = hdr(dg)
    r = find_row(dg, "liquefaction_electric 0.12")
    dg.cell(r, h["field"]).value = "liquefaction_electric (was 0.12)"
    dg.cell(r, h["rows_affected"]).value = "Parameters:liquefaction_electric"
    dg.cell(r, h["current_state"]).value = "RESOLVED"
    dg.cell(r, h["why"]).value = (
        "0.12 replaced by 0.15 on 3 Sep 2026 (Pembina Institute, Gorski and Lam 2023, "
        f"{PEMBINA}). From 27 Sep 2026 it is no longer Emission Factors range_low (now 0.26, "
        f"Delphi Group 2013, {DELPHI}); it drives the all-electric sensitivity only.")
    dg.cell(r, h["what_would_fill_it"]).value = "Closed"
    dg.cell(r, h["priority"]).value = "closed"
    r = find_row(dg, "liquefaction_electrification_assumed")
    dg.cell(r, h["priority"]).value = (
        "high - would lower liquefaction by about half at 0.15, or by over 90 percent at the "
        "0.021 grid best case, if applied to the whole slate. The all-electric sensitivity in "
        "src/drive_sensitivity.py puts a number on it.")
    print("Data Gaps: liquefaction_electric resolved; electrification priority text")

    # --- Chains: export capacity ---------------------------------------------
    ch = wb["Chains"]
    h = hdr(ch)
    r = find_row(ch, "export")
    assert ch.cell(r, h["applies_to"]).value.startswith("79.6 mtpa")
    ch.cell(r, h["applies_to"]).value = (
        "85.55 mtpa (14.0 operating, 5.85 under construction, 65.7 proposed; asset register "
        "of 27 Sep 2026)")
    print("Chains export capacity text updated")

    wb.save(path)


if __name__ == "__main__":
    main(sys.argv[1])
