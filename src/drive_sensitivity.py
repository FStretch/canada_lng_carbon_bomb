"""Liquefaction drive-type sensitivity (supplementary information only).

The central case runs gas turbine drive at 0.29 tCO2e per tonne LNG for every
terminal, for its whole operating life: the all-gas case. This sensitivity runs
the opposite end, every headline terminal on electric drive, so the two together
bracket what the choice of drive can do to the totals (decision 27 Sep 2026;
until then only Ksi Lisims and Cedar were switched).

Two electric-drive figures, both on the Parameters sheet:

    liquefaction_electric             = 0.15    Pembina Institute (Gorski and Lam
                                                2023), LNG Canada Phase 1 under
                                                electric drive with grid supply
    liquefaction_electric_drive_grid  = 0.021   BC EAO assessment of Ksi Lisims,
                                                Base Case on grid supply; the
                                                lowest sourced figure, so the
                                                best case

**Boundary caveat, carried into every output.** Both are facility intensities
(the BC EAO figure includes marine sources), not the liquefaction stage alone.
The 0.29 they replace is a liquefaction-stage factor, so the comparison is
approximate and is labelled as such.
"""

from __future__ import annotations

import pandas as pd

from src.inputs import DEFAULT_SCENARIO, get_param
from src.trajectories import _chain_intensity_split, panel_lifetime_mt

CASES = (
    (
        "all_electric_pembina",
        "liquefaction_electric",
        "Every terminal on electric drive, Pembina figure for LNG Canada (Gorski and Lam 2023)",
    ),
    (
        "all_electric_grid",
        "liquefaction_electric_drive_grid",
        "Every terminal on electric drive, BC EAO grid-supply figure for Ksi Lisims (best case)",
    ),
)
BOUNDARY_NOTE = (
    "The electric-drive figures are facility intensities (the BC EAO figure "
    "includes marine sources), not the liquefaction stage alone; 0.29 is a "
    "liquefaction-stage factor. The comparison is approximate."
)


def _row(inputs: dict, project_id: str):
    sl = inputs["assets"].loc[inputs["assets"]["project_id"] == project_id]
    if len(sl) != 1:
        raise KeyError(f"Expected exactly one register row for {project_id!r}.")
    return sl.iloc[0]


def run_drive_sensitivity(inputs: dict, panel: pd.DataFrame) -> dict:
    """Every headline terminal at each electric-drive figure, against the all-gas central."""
    params = inputs["params"]
    gas_i = float(inputs["factors"].loc["liquefaction", "central"])
    if abs(gas_i - 0.29) > 1e-12:
        raise AssertionError(
            f"Liquefaction central is {gas_i}, not 0.29. The drive sensitivity "
            "is defined against the 0.29 central and will not run if it moves."
        )
    headline_life = panel_lifetime_mt(panel, DEFAULT_SCENARIO)

    # Lifetime tonnes of LNG behind each asset's panel total, so an intensity
    # change can be applied without re-running the panel.
    meta = {}
    headline_can = 0.0
    for pid in panel.loc[panel["scenario"] == DEFAULT_SCENARIO, "project_id"].unique():
        row = _row(inputs, pid)
        total_i, _c, _h = _chain_intensity_split(
            row, DEFAULT_SCENARIO, inputs, canada_only=False
        )
        can_i, _c2, _h2 = _chain_intensity_split(
            row, DEFAULT_SCENARIO, inputs, canada_only=True
        )
        life = panel_lifetime_mt(panel, DEFAULT_SCENARIO, project_id=pid)
        meta[pid] = {"total_i": total_i, "can_i": can_i, "lng_mt": life / total_i}
        headline_can += life * can_i / total_i

    rows = [{
        "case": "central",
        "label": "Central: every terminal on gas turbine drive, 0.29",
        "liquefaction_intensity": gas_i,
        "parameter": "Emission Factors:central:liquefaction",
        "applies_to": "all terminals",
        "basis": "central case, unchanged",
        "headline_lifetime_mtco2e": headline_life,
        "headline_delta_mtco2e": 0.0,
        "headline_delta_pct": 0.0,
        "headline_canada_mtco2e": headline_can,
        "headline_canada_delta_mtco2e": 0.0,
        "boundary_note": "liquefaction stage only",
    }]
    per_asset = []
    for case, param_name, label in CASES:
        value = float(get_param(params, param_name))
        delta_i = value - gas_i
        total_delta = 0.0
        for pid, m in meta.items():
            d = m["lng_mt"] * delta_i
            total_delta += d
            per_asset.append({
                "case": case,
                "project_id": pid,
                "liquefaction_intensity": value,
                "lifetime_delta_mtco2e": d,
                "canada_delta_mtco2e": d,
            })
        rows.append({
            "case": case,
            "label": label,
            "liquefaction_intensity": value,
            "parameter": f"Parameters:{param_name}",
            "applies_to": "all headline terminals",
            "basis": label,
            "headline_lifetime_mtco2e": headline_life + total_delta,
            "headline_delta_mtco2e": total_delta,
            "headline_delta_pct": 100.0 * total_delta / headline_life,
            # Liquefaction is CAN-tagged on every chain that includes it, so
            # the whole delta lands in Canada territorial.
            "headline_canada_mtco2e": headline_can + total_delta,
            "headline_canada_delta_mtco2e": total_delta,
            "boundary_note": BOUNDARY_NOTE,
        })

    return {
        "cases": pd.DataFrame(rows),
        "by_asset": pd.DataFrame(per_asset),
        "gas_intensity": gas_i,
        "headline_lifetime_central": headline_life,
        "headline_canada_central": headline_can,
    }


def format_drive_markdown(sens: dict) -> list[str]:
    cases = sens["cases"]
    lines = []
    lines.append("## Liquefaction drive-type sensitivity (SI only)")
    lines.append("")
    lines.append(
        f"The central case puts every terminal on gas turbine drive at "
        f"**{sens['gas_intensity']:.2f}** tCO2e per tonne LNG, and the run still "
        f"asserts it. This sensitivity runs the other end: **every headline "
        f"terminal on electric drive**, at two sourced figures. **0.15** is the "
        f"Pembina Institute's figure for LNG Canada Phase 1 under electric drive "
        f"(Gorski and Lam 2023). **0.021** is the Base Case on grid supply in the "
        f"British Columbia Environmental Assessment Office's assessment of Ksi "
        f"Lisims LNG (7 August 2025, Canadian Impact Assessment Registry document "
        f"163192E, pages 847 to 848), the lowest sourced figure and so the best case."
    )
    lines.append("")
    lines.append(f"**Boundary caveat.** {BOUNDARY_NOTE}")
    lines.append("")
    lines.append(
        "| case | liquefaction tCO2e/t | headline lifetime Mt | delta Mt | delta % | "
        "headline Canada-territorial Mt | Canada delta Mt |"
    )
    lines.append("|---|---|---|---|---|---|---|")
    for _, r in cases.iterrows():
        lines.append(
            f"| {r['label']} | {r['liquefaction_intensity']:.3f} | "
            f"{r['headline_lifetime_mtco2e']:,.1f} | "
            f"{r['headline_delta_mtco2e']:+,.1f} | "
            f"{r['headline_delta_pct']:+.1f}% | "
            f"{r['headline_canada_mtco2e']:,.1f} | "
            f"{r['headline_canada_delta_mtco2e']:+,.1f} |"
        )
    lines.append("")
    lines.append(
        "Liquefaction is CAN-tagged on every chain that includes it, so the "
        "whole delta lands in the Canada-territorial column. Combustion abroad "
        "is untouched by how a terminal is powered, which is why even the best "
        "case moves the total by less than a tenth. Per-asset split: "
        "`Outputs/figure_data/sens_liquefaction_drive_by_asset.csv`."
    )
    lines.append("")
    return lines
