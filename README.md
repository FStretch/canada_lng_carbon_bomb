# Canada LNG Expansion: Lifecycle Emissions Model

An open, reproducible model of the greenhouse gas emissions caused by Canada's liquefied natural
gas export expansion, covering every LNG asset in the country in the register: operating, under
construction and proposed. Headline results are the export chain.

**Start here.** Four files carry the whole thing:

1. `README.md` — this file: method, assumptions, limitations.
2. `Inputs/Canada_LNG_Data_Inputs.xlsx` — every factor and parameter, with its type and source.
3. `Inputs/Canada_LNG_Asset_Register.xlsx` — the asset register, with a source on every cell.
4. `build_results.py` — what is asserted: every published digit is locked there.

**Status.** Submission snapshot for the manuscript *A Canadian Carbon Bomb? The potential lifetime
emissions and economic damages of Canada's LNG export expansion*, prepared for
*Environmental Research Letters*. Headline results are locked to the 29 September 2026
paper set (run 14:00 UTC). The lock includes the LNG Canada Phase 2 final investment
decision of 29 September 2026. Cite the Zenodo DOI of the GitHub release, which is the archival copy.

**Reproduce the locked results.** Python 3.13.0, then `pip install -r requirements.txt`
and `python build_results.py`.

**Manuscript map.** Figure 1 is `Outputs/figures/fig03_three_trajectories.png`.
Figure 2(a) is `Outputs/figures/fig01_stage_breakdown.png`, Figure 2(b) is
`Outputs/figures/fig02_territorial_split.png`, and Figure 2(c) is
`Outputs/figures/fig02c_territorial_trajectory.png`. Figure 3 is
`Outputs/figures/fig09_loss_damage_by_group.png`. The locked headline table is
`Outputs/paper_set_locked.csv`. One row per headline terminal is
`Outputs/si_table_assets.csv`.

## Register update, 27 September 2026

Changes to `Inputs/Canada_LNG_Asset_Register.xlsx`, each sourced in the Asset Sources sheet.
Rule adopted: where new data would move the figures, the latest published developer or
regulator figure is used, with its source and link.

| Asset | Change | Source |
|---|---|---|
| LNG Canada Phase 2 | No FID yet; status text refreshed; first LNG 2030 reconfirmed | [GEM wiki, 25 Sep 2026](https://www.gem.wiki/LNG_Canada_Terminal); [MPO](https://www.canada.ca/en/privy-council/major-projects-office/projects/national/lng-canada.html); [Northern View, 24 Sep 2026](https://thenorthernview.com/2026/09/24/lng-canada-phase-2-decision-could-come-as-early-as-october/) |
| Ksi Lisims | First exports 2029 to 2032; offtake 4 to 8 mtpa; licence early-expiry clause; grid timing | [MPO (Uniper, 29 Jul 2026)](https://www.canada.ca/en/privy-council/major-projects-office/projects/national/ksi-lisims.html); [BC, 27 May 2026](https://news.gov.bc.ca/releases/2026ECS0025-000604); [NRCan, 15 Sep 2026](https://www.canada.ca/en/natural-resources-canada/news/2026/09/canadian-lng-project-signs-third-international-offtake-deal.html); [GEM wiki](https://www.gem.wiki/Ksi_Lisims_FLNG_Terminal); [BC, NCTL](https://news.gov.bc.ca/releases/2026PREM0034-001026) |
| Fermeuse | Capacity 4.5 to 10 mtpa (developer statement); operating life 18 years, derived from the developer's 9.7 Tcf resource at 10 mtpa | [The Narwhal, 19 Mar 2026](https://thenarwhal.ca/lng-newfoundland-lessons-kitimat-bc/) |
| Cedar | Capacity 3.3 to 3.75 mtpa (BC EAO amended certificate, July 2026, same equipment); Coastal GasLink demand updated to match | [BC EAO amendment application](https://projects.eao.gov.bc.ca/api/public/document/68caf0067a03b4002205daf9/download/Cedar%20LNG%20Accommodation%20+%20Capacity%20Amendment%20Application.pdf); [GIIGNL, 22 Jul 2026](https://www.giignl.org/news/cedar-lng-receives-approval-to-increase-project-capacity); [Canadian Vanguard](https://www.thecanadianvanguard.com/cedar-lng-secures-bc-regulatory-nod-to-expand-future-output-capacity/) |
| Kino Aski | Feedgas basin now Western Canada (developer-stated); drive, cost and commercial notes | [Kino Aski LNG](https://kinoaskilng.ca/); [gasworld](https://www.gasworld.com/story/kino-aski-lng-targets-north-america-europe-energy-corridor/2257622.article/); [Federal prospectus, p. 8](https://cabinradio.ca/wp-content/uploads/2026/09/CanadaInvestmentSummit.pdf); [Energy Mix, 16 Sep 2026](https://www.theenergymix.com/kino-aski-lng-signs-tentative-deal-with-ukraine-state-fossil-as-opposition-rises-complications-abound/) |
| Kanata | Indicative terms with Hanwha; pre-FEED; FID targeted early 2028 | [LNG Industry, 14 Sep 2026](https://www.lngindustry.com/floating-lng/14092026/kanata-clean-power-and-hanwha-ocean-to-launch-pre-feed-for-proposed-floating-lng-project/) |
| Woodfibre | First exports 2028 to 2027 (in service targeted end-2027); GL-340 start deadline June 2030; 2.5 mtpa expansion noted, capacity unchanged | [LNG Journal, 16 Jul 2026](https://lngjournal.com/index.php/latest-news-mainmenu-47/item/116798-woodfibre-lng-on-track-to-complete-construction-by-late-2027); [Energetic City](https://energeticcity.ca/2025/08/29/export-licence-extension-for-woodfibre-lng-granted-by-energy-regulator/); [Federal prospectus, p. 10](https://cabinradio.ca/wp-content/uploads/2026/09/CanadaInvestmentSummit.pdf) |
| Tilbury Phase 2 | BC certificate and federal decision, 21 Sep 2026; 2.5 mtpa confirmed as an addition | [IAAC](https://www.canada.ca/en/impact-assessment-agency/news/2026/09/major-milestone-completed.html); [Surrey Now-Leader](https://surreynowleader.com/2026/09/21/delta-lng-expansion-project-issued-b-c-environmental-assessment-certificate/) |
| Tilbury Phase 1b | First exports 2028 to 2031 | [BC, 24 Jul 2026](https://news.gov.bc.ca/releases/2026ECS0043-000870) |
| Tilbury Marine Jetty | Proponent named (Tilbury Jetty Limited Partnership) | [Tilbury Pacific](https://tilburypacific.ca/about-tilbury-pacific/) |
| Kitimat LNG | Now owned by the Haisla Nation's yáqʷa Development Corporation; moved from inactive to watch (excluded from totals) | [yáqʷa, 17 Nov 2025](https://www.yaqwadevcorp.com/2025/11/18/yaq%CA%B7a-acquisition/); [BC EAO EPIC](https://projects.eao.gov.bc.ca/api/public/search?dataset=Project&keywords=Kitimat&pageNum=0&pageSize=100) |
| Placentia Bay | EA registration 2177 still active | [NL registry](https://www.gov.nl.ca/ecc/env-assessment/projects-list/) |
| Saint John | NRCan's 6.5 mtpa import capacity recorded against GEM's 7.5 | [NRCan](https://natural-resources.canada.ca/energy-sources/fossil-fuels/canadian-liquified-natural-gas-projects) |
| Tamaska | Owner now Cryopeak Energy Solutions | [Cryopeak](https://www.cryopeak.com/) |
| Northern Prince LNG (new) | Watch row: 6 mtpa developer claim, no site, filing or named buyer; only short-term CER export orders; may overlap Summit Lake | [Website](https://www.northernprincelng.com/about); [NeeStaNan](https://neestanan.ca/neestanan-and-lng-exports/); [BOE Report, 2024](https://boereport.com/2024/03/05/neestanan-signs-mou-with-northern-prince-lng/); [CER REGDOCS](https://apps.cer-rec.gc.ca/REGDOCS/Item/View/4459935) |
| Port of Churchill Plus LNG (new) | Watch row, no capacity: floating LNG concept in Hudson Bay, no developer; federal 2030 target | [MPO](https://www.canada.ca/en/privy-council/major-projects-office/projects/other/referred/churchill-plus.html); [Federal prospectus, p. 40](https://cabinradio.ca/wp-content/uploads/2026/09/CanadaInvestmentSummit.pdf); [CBC, 15 Sep 2026](https://ca.news.yahoo.com/port-churchill-railway-improvements-cost-100000703.html); [Global News](https://globalnews.ca/news/11806964/federal-government-puts-timeline-in-place-for-port-of-churchill-project-kinew/) |
| NeeStaNan Port Nelson (new) | Watch row, no capacity: feasibility study under a short-term CER export authorisation to June 2027 | [CBC](https://www.cbc.ca/1.7582120); [BOE Report, 2025](https://boereport.com/2025/07/02/neestanan-and-fox-lake-cree-nation-secure-federal-authorization-to-export-lng/) |
| Project Northern Lights (new) | Domestic small-scale plant in Alberta, shovel-ready; no capacity published, so it enters no total | [Federal prospectus, p. 7](https://cabinradio.ca/wp-content/uploads/2026/09/CanadaInvestmentSummit.pdf) |

**What the 27 September 2026 run changed.** Export capacity rose from 79.6 to 85.55 mtpa
(under construction 5.4 to 5.85; proposed 60.2 to 65.7; early_proposed 34.2 to 39.7). The
locks that moved were `EXPECTED_EXPORT_BY_CALC`, `EXPECTED_EXPORT_TOTAL`, `EXPECTED_EARLY_EXPORT`,
`EXPECTED_LIFETIME_MT`, `EXPECTED_PEAK_MT`, `EXPECTED_BUILD_OUT` and `EXPECTED_PAPER_SET_SHA256`
in `build_results.py`, and the combustion share in `LOCKED_STAGE_SHARE_PCT` in
`src/figures_report.py` (78.0 to 78.1). The territorial shares are unchanged at one decimal
(18.2 / 3.2 / 78.6).

| build-out | lifetime CO2e Mt [MC p5, p95] | lifetime CO2-only Mt | peak | ECCC 2% damages C$bn, valued when caused [MC p5, p95] | NPV to 2025 |
|---|---|---|---|---|---|
| Committed (3) | 1,917.3 [1,890, 2,094] | 1,847.2 | 59.4 in 2030 | 769 [755, 830] | 539 |
| Committed plus advanced (5) | 3,914.9 [3,817, 4,255] | 3,771.7 | 136.4 in 2037 | 1,620 [1,576, 1,742] | 1,088 |
| Full build-out (9) | 7,390.3 [6,865, 8,284] | 7,124.7 | 253.4 in 2037 | 3,176 [2,899, 3,590] | 2,015 |

At full build-out the CO2-only lifetime is 4.19% of the 170 GtCO2 remaining for 1.5°C, and the
Monte Carlo median is 7,562.4 Mt.

**Liquefaction drive, all gas against all electric** (`Outputs/figure_data/sens_liquefaction_drive.csv`,
full build-out). The central case is every terminal on gas turbine drive (0.29). The
sensitivity puts every headline terminal on electric drive:

| case | liquefaction tCO2e/t | lifetime Mt | change | Canada-territorial Mt |
|---|---|---|---|---|
| All gas turbine (central) | 0.29 | 7,556.3 | none | 1,374.5 |
| All electric, Pembina figure (Gorski and Lam 2023) | 0.15 | 7,256.1 | -300.2 (-4.0%) | 1,074.3 |
| All electric, BC EAO grid-supply figure (best case) | 0.021 | 6,979.5 | -576.8 (-7.6%) | 797.7 |

Every tonne of the change is in Canada, so the Canada-territorial total falls by 22% (Pembina
figure) to 42% (best case), while the full lifecycle total falls by 4% to 8%. Sources: [Pembina Institute, *Squaring the Circle: State of LNG 2023*](https://www.pembina.org/reports/squaring-the-circle-state-of-lng-2023.pdf);
[BC EAO, Ksi Lisims assessment report, pp. 847 to 848](https://iaac-aeic.gc.ca/050/documents/p82797/163192E.pdf). Both are facility intensities
(the BC EAO figure includes marine sources), not the liquefaction stage alone, so the
comparison with 0.29 is approximate.

**Code changes made with this update.** (1) New register field
`first_export_year_assumes_fid`. Where it is True the model does not add the five-year FID
delay (`fid_delay_mid`) to a pre-FID asset, because its `first_export_year` is a developer
date that already allows for the build after an FID. It is set for Ksi Lisims (first
deliveries 2032 under the Uniper agreement) and Tilbury Phase 1b (in service 2031). Code:
`_date_assumes_fid` and `_fid_delay_applies` in `src/model.py`, used by `src/trajectories.py`
and `src/monte_carlo.py`. (2) The Monte Carlo year range now ends at the last year any draw
can emit (sampled-life ceiling, held life, or licence stop) instead of start + 49 for every
asset, which ran past the ECCC schedule's last year (2080) once Ksi Lisims moved to 2032. On
the 3 September register both changes reproduce the locked Monte Carlo figures exactly.

Method decisions of 27 September 2026, made after the register update:

(3) **The FID delay moves a stated or derived life later instead of cutting it.** Where an
asset's life is `project_life_years` with no licence end and no authorised term, the delay now
shifts the whole life. Fermeuse runs its full 18 years (2035 to 2052, 506.9 Mt; was 13 years
and 361 Mt) and Summit Lake its full 30 (2035 to 2064, 234.5 Mt; was 194.5 Mt). The 18 years
are derived from the developer's 9.7 Tcf (275 bcm) Jeanne d'Arc resource
([Energy Mix, 5 Sep 2025](https://www.theenergymix.com/lng-developer-announces-15b-project-off-newfoundland-says-carney-policy-changes-made-it-happen/)) at the developer's 10 mtpa
([The Narwhal, 19 Mar 2026](https://thenarwhal.ca/lng-newfoundland-lessons-kitimat-bc/)), with liquefaction fuel netted:
275 / (10 × 1.38 × 1.1055) = 18.0 years. The 30 years are the developer's own
([BC EAO draft Initial Project Description, p. 10](https://www.projects.eao.gov.bc.ca/api/public/document/65cce8f768c392001b2d5692/download/Draft%20Initial%20Project%20Description.pdf): "The Project is expected to
operate for 30 years"). Licence terms, licence end years and the 40-year default keep the delay
inside the window as before. The Monte Carlo does the same draw by draw, so the operating years
stay fixed whatever delay is drawn. Code: `_life_shifts_with_delay` and `_lifespan` in
`src/model.py`; `life_shifts` in `src/monte_carlo.py`. Effect: +186.4 Mt at full build-out.

(4) **Liquefaction is sampled in the Monte Carlo** on a triangle of 0.26 / 0.29 / 0.36, the
range for gas-turbine plants in [Delphi Group (2013)](https://www2.gov.bc.ca/assets/gov/environment/climate-change/ind/lng/lng_emissions_benchmarking_-_march_2013.pdf): Sabine Pass 0.26 (p. 23,
Figure 3) and Pluto LNG 0.36, a GE Frame 7EA gas-turbine plant (p. 14), in a benchmark of
operating and proposed plants spanning 0.17 to 0.49. Emission Factors range_low moves from
0.15 to 0.26 and range_high 0.36 is now cited. Electric drive stays outside this range on
purpose. The central does not move; the Monte Carlo interval and median move slightly up.

(5) **The drive sensitivity puts every terminal on electric drive** (`src/drive_sensitivity.py`),
at 0.15 and at 0.021, against the all-gas central. It replaces the earlier two-asset version
(Ksi Lisims and Cedar only). Results are in the table above.

(6) **The Kino Aski United States supply case is dropped** from `src/feedgas_sensitivity.py`,
because the developer now names Western Canadian gas ([Kino Aski LNG](https://kinoaskilng.ca/)).
The two Western Canada pipeline-length cases (x3 and x5) stay.

(7) **Figure titles removed** from figures 1, 2(a), 2(b) and 3 (`src/figures_report.py`,
`src/loss_damage.py`), matching the manuscript, where the caption carries the title.

(8) **Inputs workbook tidy.** The liquefaction low value in the Emission Factors text and the
README sheet read 0.12 and now read correctly. The Scenarios `high_50yr` note now records that
Bill C-15 made 50 years the legal maximum CER export licence term on 26 March 2026
([Major Projects Office, LNG Canada](https://www.canada.ca/en/privy-council/major-projects-office/projects/national/lng-canada.html)). The `fid_delay_mid` note records rules (1) and
(3). The stale fig08 pointer in `build_results.py` now points to the drive sensitivity.

**Still to carry into the manuscript.** The locks, regenerated outputs and version metadata
were updated with the 29 September 2026 run. The manuscript, SI and report still need these
numbers. The manuscript's SI Monte Carlo table says liquefaction is "fixed at 0.29" and needs
the 0.26 / 0.29 / 0.36 triangle instead. LNG Canada Phase 2 took its final investment decision
on 29 September 2026; `fid_confirmed` is set and the delay no longer applies to that row.

## Register update, 29 September 2026

The LNG Canada partners (Shell, PETRONAS, PetroChina, Mitsubishi, KOGAS) took the Phase 2
final investment decision on 29 September 2026. The row leaves `advanced_proposed` for
`under_construction` (status `construction`, first export 2031). A Fluor joint venture is
the engineering, procurement and construction contractor, and Coastal GasLink Phase 2
proceeds. Committed is now four terminals. Sources are on the register row.

**What the 29 September 2026 run changed.** Export capacity stays 85.55 mtpa. Under
construction moves from 5.85 to 19.85 mtpa and proposed from 65.7 to 51.7 mtpa (advanced
26.0 to 12.0; early_proposed stays 39.7). The combustion share in `LOCKED_STAGE_SHARE_PCT`
moves from 78.1 to 78.0. The territorial shares are unchanged at one decimal
(18.2 / 3.2 / 78.6). The live paper set is in the next section.

---

## What this produces

**The paper set.** Locked 29 September 2026 (run 14:00 UTC). Central case, with the Monte
Carlo 5th to 95th percentile as its interval (10,000 draws, seed 20260828). Central is the
point estimate from the central factor values.

| build-out | lifetime CO2e Mt | lifetime CO2-only Mt | peak | ECCC 2% damages C$bn, valued when caused | ECCC 2% damages C$bn, NPV to 2025 |
|---|---|---|---|---|---|
| Committed (operating + under construction, 4 assets) | **2,967.4** [2,925, 3,241] | **2,858.9** [2,791, 3,063] | **100.9** in 2033 [99.4, 110.2] | **1,198** [1,176, 1,293] | **834** |
| Committed plus advanced (5 assets) | **4,080.8** [4,022, 4,457] | **3,931.6** [3,838, 4,212] | **136.4** in 2034 [134.5, 149.0] | **1,678** [1,647, 1,811] | **1,137** |
| Full buildout (9 projects, 85.55 mtpa) | **7,556.3** [7,076, 8,481] | **7,284.6** [6,741, 8,032] | **253.4** in 2037 [249.8, 276.4] | **3,234** [2,972, 3,659] | **2,064** |

**How the headline is reported.** As the three-point ladder above, with the tier label on each
figure; the single full-slate number is not fronted alone, since about three fifths of it has
no final investment decision. The Monte Carlo interval on each row is factor uncertainty
*conditional on that build-out*. It is not the uncertainty on what Canada's expansion will emit:
that is the spread from the committed row to the full row, about five times wider.

**Damages are two figures answering two questions.** The calendar-year sum values each year's
damage when it is caused, in constant 2025 dollars: the loss-and-damage figure, and the central.
The NPV discounts the same stream to 2025 at the same 2%: the cost-benefit figure, and the
aggregation Government of Canada regulatory guidance uses. Wherever one appears, so does the other.

The Monte Carlo median sits above the central case (7,759.6 Mt at full buildout) because the
sampled stage triangles are right-skewed, shipping 0.05 / 0.12 / 0.31 especially. It is stated
once, with that reason, and is not the reported figure. The upstream triangle high is the GWP20
scenario factor (0.440), not Howarth 0.55.

At full buildout:

| | |
|---|---|
| Headline annual (panel peak) | **253.4 MtCO2e** in 2037 |
| life_average_annual_mt | **230.0 MtCO2e/yr** (not a calendar year) |
| Lifetime emissions | **7,556.3 MtCO2e** (calendar panel 2025–2069) |
| Lifetime, CO2 only | **7,284.6 MtCO2** (plus 9,117 kt CH4) |
| Export capacity | 85.55 mtpa across nine projects |
| Scope 1 and 2 | 41.8 Mt/yr, 18.2% |
| Scope 3 | 188.2 Mt/yr, 81.8% |

Split by where the emissions are counted: **18.2% Canada, 3.2% international marine bunkers,
78.6% foreign**. The CO2-only lifetime is 4.3% of the 170 GtCO2 remaining for 1.5°C.

Every figure in the paper-set table is locked as `EXPECTED_BUILD_OUT` in `build_results.py` and
asserted on every run.

Headline results include assets whose chain is in `headline_scope_chains` (export) and whose
calc_group is in `headline_scope_calc_groups` (operating, under construction, proposed). Eight
non-export assets totalling 252.1 MtCO2e stay in the register as a stated exclusion.

The published lifetime total is the sum of a per-asset, per-calendar-year panel from 2025
through each asset's last emitting year (currently 2069). It is not duration × life-average.
The headline annual figure is the panel peak (253.4 MtCO2e in 2037). `life_average_annual_mt`
(230.0 MtCO2e/yr) is a life-average of utilisation over each facility's operating window,
including start-up years; it is not a calendar year.

These headlines are below the 9,298.1 Mt / 100.1 mtpa figures that included Discovery LNG as
early_proposed. Discovery is now cancelled, matching Global Energy Monitor. That is a scope
change, not a change in the physics of the remaining nine assets. A further 38.9 Mt came off on
3 September 2026 when regasification moved from an uncited 0.04 to the cited 0.021 of
Mukherjee et al. (2025), and a further 53.3 Mt when pipeline transport moved from an assumed 0.10 to
0.074, the figure two independent cited routes converge on. Then 55.4 Mt came back on when the
upstream factor was re-derived on that pipeline value: the old 0.22 inventory figure had netted
the retired 0.10 from the CER anchor, and netting 0.074 instead gives 0.246, a central of 0.277.
Finally 50.3 Mt came off when Fermeuse's derived capacity was netted for the gas liquefaction
burns as fuel (5.0 to 4.5 mtpa). The two 3 September corrections run in opposite directions and
net to +5.1 Mt against the 7,162.0 Mt that preceded them.

---

## Repository structure

```
Inputs/
  Canada_LNG_Asset_Register.xlsx    facts about physical assets, one row per terminal or unit
  Canada_LNG_Data_Inputs.xlsx       emission factors, parameters, scenarios and chain definitions
  loss_damage/                      SC-CO2 schedules, currency conversion, L&D parameters
  loss_damage_r1/                   Burke et al. (2026) pulse damages and the raw replication
                                    files the derived CSVs came from
src/
  inputs.py                         loads and validates both workbooks, exposes typed frames
  model.py                          the lifecycle calculation, per-gas split and aggregation
  trajectories.py                   calendar panel (published lifetime) and 2025–2050 figures
  scope.py                          headline chain/calc_group filter and build-out membership
  loss_damage.py                    global L&D, ECCC per gas + Burke upper bracket
  monte_carlo.py                    sampled physics, then ECCC pricing per gas
  lca_comparison.py                 comparator studies on their own boundaries
  benchmark_table.py                six-row boundary-aligned comparison to external estimates
  placeholder_sensitivity.py        first-export-year fill sensitivity (SI)
  lifespan_sensitivity.py           uniform 40-year life, no licence stop (SI)
  feedgas_sensitivity.py            Kino Aski Western Canadian feedgas pipeline length (SI)
  drive_sensitivity.py              every terminal on electric drive vs all gas (SI)
  si_table.py                       one row per headline asset for the SI
  figures_report.py                 manuscript figures and their CSV series
build_results.py                    orchestrates a run, asserts every lock, writes the outputs
requirements.txt                    exact pins for the environment the published run used
CITATION.cff                        citation metadata
Outputs/
  Canada_LNG_Emissions_Results.xlsx results by project, chain, stage, group and scenario
  paper_set_locked.csv              the rounded paper set; its SHA-256 is asserted on every run
  si_table_assets.csv               per-asset SI table
  benchmark_comparison.csv          the boundary-aligned external comparison
  figures/                          manuscript figures 1, 2(a), 2(b) and 3
  figure_data/                      CSV series behind each figure and SI table
```

The two input workbooks are the source of truth for emission factors, capacities, scenarios and
named parameters. The model reads them and never writes to them. Where the model needs a value
the inputs do not provide, it raises an error rather than substituting a default. A short list of
structural constants and figure-only fallbacks still lives in code; those are named in
Reproducing a run, below.

---

## Method

### The asset register

Nineteen assets in the calculation, drawn from Global Energy Monitor's Global Gas Infrastructure Tracker
(LNG Terminals, September 2025), Natural Resources Canada's project list, and Canada Energy
Regulator export licence records. Kanata LNG (June 2026) is added from proponent sources and is
not in GEM. Discovery LNG is in GEM as cancelled; the register now matches. Port of Hamilton has
no published capacity and is excluded from totals rather than estimated, as is Project Northern
Lights (added 27 September 2026). Kitimat LNG, Northern Prince LNG and the Churchill and Port
Nelson Hudson Bay concepts are watch rows, outside the calculation. Tilbury Marine Jetty
has no lifecycle chain and is excluded, not zeroed.

**Inclusion test for proposed projects.** A project enters the calculation as early_proposed
only if it has at least one of: a regulatory filing (an environmental assessment registration,
a project description or a CER licence application), a named counterparty (a buyer, partner or
equipment supplier under a signed agreement) or a named site. A project that fails the test,
or is shelved, is a **watch row** (`calc_group` = watch): it is recorded with its sources and
excluded from every total, so it can be promoted when evidence appears. A project with no
published capacity is excluded from totals whatever its status.

Three rules govern the register:

- **No fabricated sources.** Every source is a dataset, publication, regulator or filing, never a
  file in this repository. Where a value was entered manually with nothing external behind it, the
  source field says so.
- **Blank beats wrong.** Where a value is unknown the cell is empty and the source reads
  "not available". Assets without a capacity are excluded from all totals rather than estimated.
- **Capacity is the latest design figure.** Where sources disagree, the most recent design
  capacity published by the developer or permitted by a regulator is used and the alternatives
  are recorded. Cedar LNG is the worked example: its final investment decision announcement
  stated 3.3 mtpa and NRCan and the CER publish 3.0, but the BC EAO's amended certificate of
  July 2026 permits 3.75 mtpa from the same equipment, so the register carries 3.75.

Assets carry two classifications. `tier` describes progress as an export project. `calc_group`
describes simply whether the asset is running, being built, or proposed, and is what the model
groups by.

### Four lifecycle chains

Not every asset passes through the same stages. **Stages absent from a chain are not computed at
all, rather than set to zero.**

| Chain | Stages | Reason |
|---|---|---|
| Export | upstream, pipeline, liquefaction, shipping, regasification, combustion | The full chain |
| Bunkering | upstream, pipeline, liquefaction, combustion | Burned by the vessel that collected it, so never shipped as cargo or regasified |
| Domestic | upstream, pipeline, liquefaction, regasification, combustion | Regasified and burned in Canada, never shipped |
| Import | regasification, combustion | Upstream, liquefaction and shipping occurred abroad |

### Territorial tagging

Every stage carries a tag recording where the emission physically occurs:

- `CAN`, Canada territorial: counts against Canada's national inventory and targets
- `BUNK`, international marine bunkers: reported separately under UNFCCC guidelines and
  attributed to no country. This covers LNG carrier voyages as well as fuel sold to ships at a
  Canadian port
- `FOR`, foreign territorial: counted by the importing country

Territorial tagging is independent of the scope 1, 2 and 3 classification. Scope describes who
controls an emission; territory describes where it occurs. For export LNG the two nearly coincide.
For domestic LNG they diverge, because scope 3 combustion happens in Canada.

**Attributing the whole chain to the Canadian buildout is an accounting choice, not a physical
fact, and it is made deliberately.** The territorial tags record where each tonne is emitted; the
lifetime total then attributes every tonne, including the 79 per cent combusted abroad, to the
Canadian decision to build. That is extraction-based accounting, one of the three recognised
frames alongside territorial (production) and consumption accounting: Davis, Peters and Caldeira
(2011) formalise it as the supply chain of CO2 traced back to the fossil fuel extracted; Andrew,
Davis and Peters (2013) show how the three frames allocate the same emissions differently; and
Steininger et al. (2016) argue the frames should be used together because each makes a different
actor's responsibility visible. This model reports the territorial split *and* the extraction
total so a reader can use either.

### Emission factors

Each stage carries a single central value in tonnes of CO2 equivalent per tonne of LNG. Low and
high figures are not used to construct scenarios, because the stages are not correlated; they set
the triangle each stage is sampled on in the Monte Carlo.

| Stage | Central | Basis |
|---|---|---|
| Upstream production | 0.277 | Canada Energy Regulator British Columbia oil and gas emissions, net of the 0.074 pipeline stage, corrected for measured methane on the methane share; every step in the basis cell |
| Pipeline transport | 0.074 | Liu et al. (2021) via the ERA/Modern West restatement, 0.0735; cross-checked against CER/ECCC 2019 national pipeline transport, 0.0744. Applied flat, not distance-scaled |
| Liquefaction | 0.29 | Gas turbine drive, applied to every terminal for its whole operating life. Range 0.26–0.36, the gas-turbine plants in [Delphi Group (2013)](https://www2.gov.bc.ca/assets/gov/environment/climate-change/ind/lng/lng_emissions_benchmarking_-_march_2013.pdf) (Sabine Pass to Pluto LNG). Electric drive is a separate sensitivity, not part of the range |
| Shipping | 0.12 | Howarth (2024), reconstructed against IMO carrier data. Derived on the British Columbia to north-east Asia route and **scaled per asset by route distance** |
| Regasification | 0.021 | Mukherjee et al. (2025), *Communications Earth & Environment* 6:16, peer-reviewed US LNG lifecycle assessment. Range 0.011–0.0275 from IEA (2025) |
| Combustion | 2.75 | Stoichiometric methane (44/16), the right central for de-inerted LNG. Range 2.58–3.00 derived from the IPCC 2006 uncertainty intervals |

**Shipping is scaled by route distance per asset.** The 0.12 central factor is derived on the
British Columbia to north-east Asia route (`route_bc_to_northeast_asia_nm` = 3,800 nm). Each
asset's shipping intensity is `0.12 × route_distance_nm / 3,800`. Seven British Columbia
terminals sit at the 3,800 nm basis; Kino Aski is 2,980 nm and Fermeuse 2,470 nm to Rotterdam,
so the two Atlantic projects carry proportionally less shipping. An asset on a chain that
includes the shipping stage with a blank `route_distance_nm` raises rather than defaulting to
the BC basis; bunkering has no shipping stage. The lifecycle-intensity
comparison in `src/lca_comparison.py` deliberately stays on the unscaled BC basis, because the
comparator studies are single-route figures.

**Canadian sources are used wherever the stage occurs in Canada.** Upstream, pipeline and
liquefaction all happen here and draw on Canadian regulatory and inventory data. Shipping,
regasification and combustion occur elsewhere and use international sources. That boundary
matches the scope 1 and 2 versus scope 3 split.

**Two factors moved onto cited values on 3 September 2026.** Regasification was an uncited 0.04
attributed to the RMI Oil Climate Index, which states no such figure; it is now **0.021** from
Mukherjee et al. (2025), with a 0.011–0.0275 range from the IEA's 0.2–0.5 gCO2e/MJ (converted at the
IEA's own 55 MJ/kg basis). That took 38.9 Mt off the lifetime. The combustion **range** was an
uncited 2.50–3.00; it is now **2.58–3.00**, derived from the IPCC 2006 Vol. 2 Ch. 1 uncertainty
intervals on natural gas NCV (Table 1.2) and carbon content (Table 1.3), applied as a relative
band to our 2.75 central. The central 2.75 is unchanged: it is the stoichiometric value for pure
methane, which is right for LNG because liquefaction strips inerts and heavier fractions. IPCC's
own natural-gas central of 2.693 is for pipeline gas.

**Pipeline transport moved onto two converging cited routes on 3 September 2026**, from an
assumed 0.10 to **0.074**. Route A: Liu et al. (2021) report 4.2 gCO2e/MJ at transmission
pipeline outlet, and the ERA / Modern West restatement puts the same study at 2.86 gCO2e/MJ at
plant exit; the difference, 1.34 gCO2e/MJ, is the transmission stage, which at 54.863 GJ/t is
0.0735 tCO2e/t. Route B: the CER reports 8.3 MtCO2e of Canadian pipeline transport emissions in
2019, primarily compressor-station combustion, which over roughly 170 bcm of marketable gas at
36 PJ/bcm is 0.0744 tCO2e/t. The two agree to within 1.2%; the retired 0.10 was about 35% above
both.

**The pipeline factor is applied flat, not scaled by distance** — unlike shipping. Route A is
calibrated on a 1,100 km line and Route B is a national average haul, while the in-scope feedgas
lines are Coastal GasLink 670 km, Prince Rupert Gas Transmission 750 km and FortisBC Eagle
Mountain 47 km. If transmission emissions scale with distance, 0.074 is high for all three,
Woodfibre most of all. That is a stated limitation rather than a correction applied here. The
0.037–0.133 range is a declared assumption: it keeps the relative width of the retired band, and
the 1.2% agreement between the two routes is convergence on the central, not an uncertainty
interval.

**The pipeline factor sets aside British Columbia evidence, deliberately.** The model's rule is
that Canadian sources are used wherever a stage occurs in Canada, and for this stage a British
Columbia document exists: the Coastal GasLink assessment implies about 0.094 tCO2e/t at design
capacity, above the retired 0.10-era check and the adopted 0.074 alike. It was not adopted
because it is a projection from an environmental assessment, made before the line operated,
whereas Liu et al. (2021) is peer-reviewed and measurement-informed and the CER route is an
observed national inventory; two independent routes agreeing to 1.2 per cent is the stronger
evidence, and the projection sits inside the sampled band. The basis cell says the same.

All the external comparisons the model holds are now assembled in one place — see **Benchmark
comparison** in `Outputs/benchmark_comparison.csv`. Every
comparison that is not an adopted value points the same way: this model sits at or below the
external figure.

**Electric drive is not assumed for any terminal.** Of the projects claiming electrification, only
Cedar and Woodfibre have interconnection works under construction; the remainder rest on memoranda
of understanding, aspirations conditional on transmission that has not been built, or statements
with no firm date. Ksi Lisims holds an MOU with BC Hydro for up to 600 MW conditional on the North
Coast Transmission Line expansion, which is not built; its second phase is due by winter 2032. LNG Canada Phase 2 states an intention to
transition to electric motors as more renewable power becomes available, with no date and no
contracted supply. A commitment contingent on infrastructure that does not exist is not a basis for
lowering the emissions factor. Previous drive classifications are retained in
`liquefaction_drive_note`. Where electrification is later contracted and built, this assumption
should be revisited. The drive sensitivity (`src/drive_sensitivity.py`,
`Outputs/figure_data/sens_liquefaction_drive.csv`) shows the other end: every headline terminal
on electric drive, at 0.15 and at the BC EAO's 0.021 grid-supply figure for Ksi Lisims as a best
case, so all gas and all electric bracket what the drive choice can do. That 0.15 is Pembina's
published figure for LNG Canada Phase 1 under electric drive (*Squaring the Circle: State of LNG
2023*; Gorski and Lam 2023). It
replaced 0.12 on 3 September 2026: no Pembina document states 0.12, which appears to have been a
misreading of the 0.11 tCO2e/t *reduction* in *Wellhead to Waterline* (2014) as an absolute
intensity.

**Upstream is derived rather than adopted, and every step is shown.** The CER reports British
Columbia oil and gas production, processing and transmission emissions of 14.6 MtCO2e for 2022,
against roughly 63 billion cubic metres of marketable gas. At 1.38 bcm per Mt of LNG that is
45.65 Mt LNG-equivalent, so the anchor is 0.320 tCO2e per tonne *including* transmission.
Transmission is this model's separate pipeline stage, so its central of 0.074 is removed:
0.320 − 0.074 = **0.246**, the official-inventory upstream factor. Peer-reviewed measurement
studies consistently find Canadian inventories understate upstream methane by 1.5 to 1.7 times.
Applied to the methane portion only, at the 0.25 methane share, that gives
0.246 × (1 − 0.25 + 1.5 × 0.25) = 0.2765, stored as **0.277**. The basis cell on the Emission
Factors sheet carries the same arithmetic, so the number reproduces from the workbook alone.
Until 3 September 2026 the inventory figure was 0.22, which had netted the *retired* pipeline
value of 0.10; when pipeline moved to 0.074 that arithmetic stopped closing and the chain was
dropping about 0.026 tCO2e/t against its own anchor. Re-deriving added 55.4 Mt to the lifetime.

Howarth (2024) is retained as the upper bound at 0.55 rather than used as the central value. That
study addresses United States supply chains using higher leakage assumptions, while Canadian
measurement evidence finds the Montney formation among the lowest-intensity gas plays in North
America.

### Throughput and timing

| Assumption | Value | Basis |
|---|---|---|
| Utilisation, year 1 and 2 | 40% / 70% | Model assumption |
| Utilisation, steady state | 83.9% | IGU World LNG Report 2026, global average for 2025 |
| LNG Canada Phase 1 ramp | 25 / 60 / 85% | Observed startup, first cargo June 2025 |
| Saint John import | 2.5% flat | Repsol annual report, 8 TBtu regasified in 2023 |
| Delay where no FID | 5 years | IEA Global LNG Capacity Tracker and Rogers (2017), OIES Energy Insight 4 p. 1. Tested across an asymmetric 4 to 8 years |
| Operating life | licence end, else stated or derived life, else 40 years | CER export licence expiry where filed; not a uniform 40-year run |
| Methane GWP100 / GWP20 | 29.8 / 82.5 | IPCC AR6 Working Group I, fossil methane |
| Remaining 1.5°C budget (50%) | 170 GtCO2 | Global Carbon Budget 2025, from start of 2026 |
| Remaining 1.7°C / 2°C (50%) | 525 / 1,055 GtCO2 | Same source |

**Lifespans are capped at each project's export licence expiry** rather than running a uniform 40
years. The end year may emit; the year after may not. That cut is 1.8 GtCO2e against the same
nine assets at a uniform 40 years from first export with no licence stop (9,362.7 Mt versus
7,556.3 Mt). Anyone comparing versions should treat that as a methodological improvement, not
an error in the lower total. Where a proponent states a different operating life, that is used
instead: Summit Lake PG LNG states 30 years. Fermeuse's 18 years are derived from the
developer's resource claim (see Limitations).

Lifetime emissions are also reported as a share of those remaining budgets. The budgets are
CO2, so the **paper value is the CO2-only lifetime** (7,284.6 MtCO2 at full buildout, 4.3% of
the 170 GtCO2 remaining for 1.5°C). That is like for like. The panel carries an explicit
per-gas split — `co2_mt`, `ch4_derived_co2e_mt` and `ch4_mass_kt` per asset-year, summing to
`emissions_mtco2e` exactly — built from parameters already on the workbook: upstream CH4 is
the scenario factor less the CO2 part of the official inventory, and shipping CH4 is the
measured methane-slip share of the carrier uplift (1 − 1/1.44). Two residual caveats remain:
pipeline fugitive methane is not split, so a small amount of methane sits inside the CO2
total; and non-CO2 gases other than methane (N2O, refrigerants) are not counted anywhere in
the model. The GWP100 CO2e share against the same budgets (4.3% for 1.5°C) is still reported
alongside for continuity.

The FID delay sits inside the lifespan window rather than extending it, so a delayed project has
fewer operating years. It is the least evidenced parameter in the model. Since 27 September 2026 the delay is not
added where the register marks `first_export_year_assumes_fid`: a developer date that already
allows for the build after an FID (Ksi Lisims 2032, Tilbury Phase 1b 2031). Also since
27 September 2026, where the life is a stated or derived operating life
(`project_life_years`) with no licence end or authorised term, the delay moves the whole life
later instead of cutting it: Fermeuse and Summit Lake keep all their operating years.

### Loss and damage

A separate module (`src/loss_damage.py`) multiplies the same full-lifecycle, year-of-emission
CO2e series by a social cost of carbon. It does not change the emissions totals.

The **central case** is named in `Inputs/loss_damage/parameters.csv`
(`central_price_family=eccc`, `central_aggregation=calendar_year`,
`eccc_central_discount_rate_pct=2`): ECCC official SC-CO2 **and SC-CH4**
applied per calendar year of emissions to the panel's per-gas split, converted
to 2025 CAD. ECCC 1.5% and 2.5% are the central case's sensitivity range.

**Why the calendar-year sum is the central, and why the NPV travels with it.** Each year's
social cost is already the present value, at that year, of the future damage stream from a tonne
emitted then. Summing the years without further discounting therefore values each year's damage
*when it is caused*, held at 2025 prices. That is the loss-and-damage question — what harm does
this buildout do, valued as it does it — and it is how Burke et al. (2026) aggregate a multi-year
stream. Discounting each year's damage back to 2025 at the same 2% answers a different question,
the cost-benefit one — what is the stream worth today, to weigh against benefits also expressed
today — and it is how the Government of Canada's regulatory guidance aggregates a multi-year
stream. At full buildout the two are **C$3,234 bn** valued when caused and **C$2,064 bn** as an
NPV to 2025. The paper asks the loss-and-damage question, so the calendar sum is the central; the
NPV is not a sensitivity on it but the answer to the other question, and the two are carried
together wherever the total is stated. An earlier note that "the official schedule already embeds
Ramsey discounting, so a second NPV is a sensitivity only" was true of a single year's social
cost and did not settle the cross-year question; it has been withdrawn.
Burke et al. (2026) is an upper-bracket sensitivity across discount rates and
Figure 2e horizons (default g = 0; Hatton +2% is not used). Canada's 0.17%
share of a 1990 pulse (future window) is applied to Burke damages only, and is
reported with `P_dam_FD` = 0.41 attached: only 41% of Burke draws put Canada in
net loss from that pulse at all, against 0.98 for the United States and China.
The replication package ships no country-level damages table by pulse year, so
no 2020-pulse share is available. The Conference
Board whole-chain GDP figure (Table 1: C$11.153bn/yr in 2020 CAD at 56 mtpa),
scaled linearly on proposed export nameplate (now 51.7 mtpa) and inflated to
2025 CAD, is the Canada denominator.

Damages are priced **per gas**. Each calendar year contributes
`co2_t × SC-CO2_t + ch4_mass_t × SC-CH4_t`, both official ECCC schedules, both
at the same discount rate, both inflated CAD 2021 to CAD 2025 exactly once.
The CO2 and CH4 series come from the panel's explicit split (upstream excess
over inventory CO2, plus shipping methane slip). Methane is about 1.7% of the
central damage bill. Burke has no SC-CH4, so the Burke family still prices the
whole GWP100 CO2e total at Burke's SC-CO2; that is stated in its labels.
Pricing the whole CO2e at SC-CO2, the retired treatment, would raise the
central figure by about 2%; that is reported as a one-line reconciliation.

### Scenarios

Five scenarios vary the upstream factor, each named for its basis rather than described as a
percentage.

| Scenario | Upstream | Basis |
|---|---|---|
| `inventory_as_reported` | 0.246 | Official Canadian inventory at face value, net of the 0.074 pipeline stage |
| `measurement_central` | 0.277 | Default. Three measurement studies agree on a 1.5x correction |
| `measurement_high` | 0.289 | 1.7x, from a British Columbia aircraft survey |
| `near_term_methane_gwp20` | 0.440 | Methane weighted over 20 years rather than 100 |
| `howarth_high` | 0.55 | United States focused study, high leakage assumptions |

The 20-year warming uplift is applied to the methane portion of upstream emissions only. The
non-methane portion of the inventory factor is unchanged. The methane portion is multiplied by
the 1.5 measurement correction and then by the ratio of GWP20 to GWP100 for fossil methane
(82.5 / 29.8). The result, 0.440, depends on the methane share of 0.25 and the 0.246 inventory
base. Applying the
uplift to the whole factor would assume all upstream emissions are methane. Pipeline methane is
not re-weighted: no pipeline methane share is defined in the workbook.

### Declared assumptions

These are judgements, not derivations. They are typed as such on the Parameters or Emission
Factors sheets.

- The 2030 overshoot gap of 200 Mt is a round figure. The interval behind it is 191 to 229 Mt
  (projected 646 Mt minus the 417–455 Mt target range).
- Upstream factors are stored to three decimals: 0.277 and 0.289 are rounded from 0.2765 and
  0.2888 on the 0.2458 inventory base. Two-decimal storage was dropped on 3 September 2026
  because it hid the dependence of the stored central on the methane share.
- Shipping is held at 0.12 although the IMO reconstruction gives 0.110.
- The Conference Board GDP figure is scaled linearly from 56 mtpa to the proposed export
  nameplate (51.7 mtpa).
- First-export years of 2033 and 2035 are placeholder sensitivity cases, not filed dates.
  Central fill remains 2030.
- The SI drive sensitivity puts every headline terminal on electric drive. It is a bound, not a
  forecast: the central case is gas turbine for every terminal.
- The pipeline 0.037–0.133 range is assumed. The central 0.074 is cited; no published
  uncertainty interval for Canadian gas transmission intensity was located.
- The liquefaction Monte Carlo range, 0.26 to 0.36, is the gas-turbine span of one 2013
  benchmark ([Delphi Group 2013](https://www2.gov.bc.ca/assets/gov/environment/climate-change/ind/lng/lng_emissions_benchmarking_-_march_2013.pdf), Sabine Pass to Pluto LNG). Electric drive is kept
  out of it on purpose and run as its own sensitivity.
- `upstream_ch4_share` is **0.25**, the centre of the 18–34% bracket that Johnson et al. (2023)
  supports. It is still a judgement within a bracket rather than a single reported figure, so it
  stays typed as an assumption.
- The FID-delay band **4 / 5 / 8** is deliberately asymmetric. The cited mid of 5 years is
  "typically 5 years, *before* unforeseen slippage", and slippage in LNG construction runs one
  way. A symmetric band would imply a project is as likely to be early as late.

### Proponent-sourced values

These are taken from the proponent, not from an independent regulator or inventory, and are
labelled as such in the register.

- Cedar LNG at 3.75 mtpa, from the developer's certificate amendment approved by the BC EAO in
  July 2026. The FID announcement said 3.3; NRCan and the CER publish 3.0.
- Fermeuse at 10 mtpa (developer statement to The Narwhal, March 2026) and its 18-year life,
  derived from the developer's 9.7 Tcf resource.
- Ksi Lisims first exports in 2032 and Woodfibre in 2027 (developer statements, July 2026).
- Kino Aski at 15 mtpa (proponent press coverage; no regulatory process has begun).
- Kanata LNG at 12 mtpa (proponent; not in GEM).
- Summit Lake's 30-year operating life.
- LNG Canada Phase 1 ramp 25 / 60 / 85%.
- The Conference Board GDP figure, published via the LNG Alliance.
- The 3,800 nautical mile British Columbia to north-east Asia route, via CAPP / Oxford.

---

## Validation

**Pipeline transport.** The British Columbia Environmental Assessment Office's 2014 assessment
of Coastal GasLink put operations at 3.517 MtCO2e a year at the line's full 5 Bcf/d design
capacity. Over that throughput (51.7 bcm a year, 37.4 Mt of LNG-equivalent at 1.38 bcm per Mt)
it implies about **0.094 tCO2e per tonne**. The model holds **0.074**, from two independent
routes that agree to 1.2 per cent: Liu et al. (2021) on an 1,100 km Alberta line and the
CER/ECCC national pipeline-transport inventory. The assessment figure is 27 per cent above the
adopted value. It is a pre-operation projection rather than a measurement, which is why the
peer-reviewed and inventory routes are preferred; the gap is inside the sampled 0.037–0.133
band, and if the line runs at its assessed intensity the model is low on the LNG Canada and
Cedar share of throughput by roughly 1 MtCO2e a year. An earlier version of this paragraph
reported a 0.6 per cent agreement between the model and the assessment; that was the output at
the retired assumed value of 0.10 and confirmed only the arithmetic of scaling it.

**Shipping.** The 0.12 factor was reconstructed from the International Maritime Organization's 2023
LNG carrier fleet average of 8.50 gCO2 per deadweight tonne per nautical mile, applied over a round
trip and uplifted for measured methane slip, giving 0.110. Within 8 per cent. The model holds
0.12; 0.110 is the reconstruction.

**Liquefaction.** LNG Canada Phase 1 produces 4.06 MtCO2e a year on a nameplate basis, against the
approximately 4 Mt in the facility's own environmental assessment.

Every run also asserts that capacity totals tie to the register, that scope and territorial splits
sum to the total, and that no stage is computed for a chain that does not include it.

---

## Limitations

**No displacement counterfactual is modelled, and that is a choice.** The quantity reported is
the gross emissions committed by the buildout on the extraction axis: every tonne of LNG the
slate would produce, followed through to combustion. The model does not ask whether that gas
displaces coal, other LNG, or nothing, and does not net any of it off. Gross extraction-basis
accounting is a recognised frame with regulatory precedent. The UK Supreme Court in *R (Finch) v
Surrey County Council* [2024] UKSC 20 held that downstream combustion emissions are an effect of
an extraction project that must be assessed; the UK government's supplementary guidance that
implements it (Department for Energy Security and Net Zero, *Environmental Impact Assessment:
assessing effects of downstream scope 3 emissions on climate*, June 2025) directs that the
starting point is "a (rebuttable) presumption that all produced hydrocarbons over the lifetime of
a project will eventually be combusted" (p. 9), and that substitution "is not considered to be a
factor affecting whether scope 3 emissions from a project's downstream activities are an effect
that needs to be assessed" (p. 7). That is exactly this model's construction. A displacement
argument is a separate claim that a proponent can make with evidence; it does not change what
the buildout commits.

**Construction emissions are excluded.** Literature suggests these would add roughly 50 to 150
MtCO2e, about 1 to 2 per cent of the lifecycle total. All figures are conservative in this respect.

**Abandoned well methane is not counted.** Measured diffusive flux from abandoned shale wells
reaches 20,000 cubic metres per well per day. The buildout requires 12,558 to 32,000 new wells, all
of which are eventually abandoned.

**The methane share of upstream emissions is a bracketed judgement, not a reported figure.**
British Columbia does not publish the carbon dioxide and methane split separately. The share is
**0.25**, set on 3 September 2026 from Johnson et al. (2023): their British Columbia 2021 upstream
methane intensity of 0.38 per cent of marketed gas, divided by the 1.7 times factor by which their
measurement-based inventory exceeds the official one, implies a methane share of 25 to 30 per cent
depending on the methane GWP vintage; an absolute route over the same paper's 144.5 kt/y gives 21
to 25 per cent. The bracket spans roughly 18 to 34 per cent and centres near 25, which is the value
adopted. It replaced an assumed 0.30 that sat high in the bracket without a stated reason. The
share now visibly moves the stored central: 0.246 × (1 − s + 1.5s) is 0.283 at a share of 0.30 and
0.277 at 0.25, stored to three decimals since 3 September 2026. Under the earlier two-decimal
storage both rounded to 0.25 and the headline looked insensitive to the share; that was a rounding
artefact, not a property of the model. The share also sets the CO2/CH4 split, the SC-CH4 damages
line, and the `near_term_methane_gwp20` scenario.

**Three projects rest on weak capacity figures.** Kino Aski LNG's 15 mtpa (formerly Marinvest,
Baie-Comeau) is from a 17 August 2026 press release; no regulatory process has begun and the
feedgas pipeline route is undefined, though the developer now names Western Canadian gas.
Fermeuse's 10 mtpa is a developer statement to journalists, with no filing. Summit Lake's 2.7
mtpa is an upper bound from an assessment the proponent asked to suspend. All three are
identified by tier and can be removed from any total.

**Discovery LNG is cancelled in this version.** GEM records cancelled (inferred 4 y). No
regulatory filing, proponent statement or news of current activity was found after GEM's
September 2025 snapshot. The register had carried it as early_proposed at 20 mtpa on the basis
of a third-party table that was wrong about Grassy Point, Bear Head and Goldboro. It is now
inactive, matching GEM. That removes 2,043.9 MtCO2e and 20 mtpa from the headline.

**Fermeuse's 10 mtpa is a developer statement, and its life is derived.** Crown LNG's chief
executive told The Narwhal (19 March 2026) the plant would export up to 10 mtpa through a 380 km
subsea pipeline; no LNG plant application has been filed. The register carries 10 mtpa from
27 September 2026. The developer's own resource claim, 9.7 Tcf (275 bcm) in the Jeanne d'Arc
Basin, then bounds the operating life: at 10 mtpa, with the gas liquefaction burns as fuel
netted at the model's own factors (1.1055 t of feedgas per t of LNG), the resource lasts
275 / (10 × 1.38 × 1.1055) = 18 years at full output, recorded in `project_life_years`. Running
10 mtpa for 40 years would need about 21.5 Tcf, twice the developer's claim. Until 27 September
2026 the register instead derived capacity from the resource over a 40-year life (4.5 mtpa from
3 September, 5.0 before that); that derivation is superseded.

**Tilbury Phase 2's classification is a judgement.** FortisBC describes the expansion as serving
Lower Mainland resilience and marine fuelling; NRCan lists it as an export project. This analysis
places it on the bunkering chain and records the disagreement rather than resolving it.

**Capacity is never summed across chains.** Export and bunkering capacity is liquefaction; import
capacity is regasification. They measure opposite operations.

**The LNG Canada Phase 1 ramp has the right average and the wrong shape.** The model assumes
25 / 60 / 85 per cent for years one, two and steady state. Observed throughput now exists: first
cargo 30 June 2025, Train 2 in production from November 2025, roughly 4.6 Mt exported to
mid-March 2026, the first month above 1 Mt in April 2026 and 1.2 Mt a month by May 2026. Against
14 mtpa nameplate that is about 11 per cent of a full year in calendar 2025 against the assumed 25
per cent, then 86 and 103 per cent of nameplate in April and May 2026 against the assumed 60 per
cent for year two. The model starts too high and ramps too slowly, and the two errors largely
offset over the asset's life. Recorded against the parameters rather than corrected, because one
facility's observed profile is not a basis for the generic ramp applied to every other asset.
The same shape problem applies to the generic 0.40 / 0.70 ramp on the other eight assets: the
observed data say it is too high in year one and too slow to full rates. Lifetime totals are
insensitive, because the two errors offset; the **2037 peak year is not**, because the peak is
set by which assets reach full rates in which calendar year. **No ramp-shape sensitivity
exists.** A reader should treat the peak year as ramp-dependent to within a year or two either
side, and the peak magnitude as more robust than its timing.

**The chain sits at or below every external benchmark it is compared against, and here is what
that is worth.** `Outputs/benchmark_comparison.csv` assembles six boundary-aligned comparisons;
none shows this model above the external figure. Liquefaction is held at 0.29 against the IEA's
0.33 global average: adopting the IEA figure would add roughly **84 Mt** to the
7,556 Mt lifetime. Shipping is held at 0.12 against the IEA's 0.18 to China: adopting that
figure as it stands would add roughly **120 Mt**, and on a per-kilometre basis, since the
IEA voyage is shorter, the IEA intensity would add closer to **300 Mt**. On the aligned
well-to-regasification boundary the model is 0.78 tCO2e/t against Roman-White et al.'s expected
1.19, and on liquefaction plus shipping plus regasification it is 0.43 against Balcombe et al.'s
0.62 to 1.71. Every held-low choice is recorded with its reason in the Emission Factors sheet;
the direction is uniformly conservative, and a reader who prefers the external values can scale
from the stage series in `Outputs/figure_data/fig01_stage_breakdown.csv`.

**Annual figures come in two forms.** The headline annual is the panel peak (253.4 MtCO2e in
2037). `life_average_annual_mt` (230.0 MtCO2e/yr) is a life-average across each facility's
operating window. The published lifetime (7,556.3 MtCO2e) is the sum of the calendar panel
from 2025 through 2069. Duration × life-average is no longer published.

**One route distance is the west-coast basis rather than a port-specific figure.** Kanata LNG
(Prince Rupert) has no published port-to-port sailing distance, so it takes the cited west-coast
Canada figure of 3,800 nm used for the other BC export assets. Prince Rupert is nearer north Asia,
so that is a small overstatement. Recorded in the register's Data Gaps sheet.

**The oil comparison has been retired.** Figure 7 compared Canadian LNG against Trans Mountain
and a proposed Alberta–BC bitumen line. It was removed from the model on 3 September 2026 when
the Trans Mountain slide was dropped from the deck. It rested on three addends (0.075 upstream,
0.008 transport, 0.43 combustion) that were never separately sourced, applied NRCan's oil-sands
intensity to a mixed ticket slate, ran oil at nameplate × 365 days × 40 years against an LNG side
that carries a ramp, a utilisation curve and an FID delay — understating LNG by roughly 19% — and
used an unvalidated 1 Mbpd capacity for a pipeline with no published design figure. Rather than
leave it generating an unpublished figure on unsourced arithmetic, the figure, slide table 15,
`src/report_params.py` and the seven `tmx_*` / `days_per_year` parameters were all removed. The
history is in git.

---

## Reproducing a run

Built and published on **Python 3.13.0** (Windows 11). Dependencies are pinned exactly in
`requirements.txt` — `pandas==2.3.2`, `numpy==2.3.3`, `matplotlib==3.10.8`, `openpyxl==3.1.5` —
rather than ranged, because the run asserts a SHA-256 over the locked paper-set table and a
library upgrade that changed float formatting should fail the run rather than quietly republish
different digits.

```bash
pip install -r requirements.txt
```

```bash
python build_results.py
```

The build sets Matplotlib's Agg backend, so figures write without a display.

Reads both workbooks from `Inputs/`, runs all five scenarios, and writes the results workbook,
the locked paper set, the SI and benchmark tables, and the figures to `Outputs/`. The inputs are opened read-only
and the run asserts they are unmodified on completion.

**Every published number is locked.** `EXPECTED_BUILD_OUT` in `build_results.py` holds the paper
set for all three build-outs, and `EXPECTED_PAPER_SET_SHA256` is a SHA-256 over
`Outputs/paper_set_locked.csv`, the rounded canonical copy of that table. A re-run that moves any
published digit fails with the expected and actual hash rather than silently rewriting the
outputs. Re-locking is deliberate: both constants change in the same commit, with the old and
new values in the commit message.

Repository metadata for citation is in `CITATION.cff`.

Workbook edits are deliberate and scripted. The model never writes to `Inputs/`; the one-off
scripts that changed an input workbook or the register live in `tools/`, each documenting what it
changed and why. The 27 September 2026 edits are
`tools/register_update_2026-09-27.py`, `tools/register_update_2026-09-27b.py`,
`tools/register_update_2026-09-27c.py` and `tools/data_inputs_update_2026-09-27.py`.
They are records of edits already applied. Do not run them again.

To change an assumption, edit the input workbooks rather than the code. Emission factors,
capacities, scenarios, utilisation, GWP values, FID and life bounds, and the LCA comparator
totals now live on those sheets.

**What remains in code, and why.** Unit conversions and plot styling. The 2025–2050 figure-year
window. Build-out membership (which tiers sit in committed). Assertions that lock published
digits so a drifted workbook fails the run. The Kino Aski
SI distance multipliers (×3 and ×5) and the 670 km Coastal GasLink length used only in that
sensitivity. None of those is an emission factor or a project capacity.

---

## Principal data sources

**Infrastructure**
Global Energy Monitor, Global Gas Infrastructure Tracker: LNG Terminals (September 2025) and Gas
Pipelines (November 2025); Global Oil and Gas Extraction Tracker (March 2026). Tracker data is
licensed CC BY 4.0; attribution is required.

**Government and regulatory**
- *R (Finch) v Surrey County Council* [2024] UKSC 20. https://www.supremecourt.uk/cases/uksc-2022-0064
- Department for Energy Security and Net Zero (June 2025). Environmental Impact Assessment:
  assessing effects of downstream scope 3 emissions on climate. Supplementary guidance for
  offshore oil and gas projects.
  https://assets.publishing.service.gov.uk/media/6853fa3d1203c00468ba2b15/Supplementary_guidance_-_Effects_of_Scope_3_Emissions.pdf
Canada Energy Regulator: provincial and territorial energy profiles, export licence applications,
Canada's Energy Future. Environment and Climate Change Canada: National Inventory Report. Natural
Resources Canada: Canadian LNG projects. British Columbia Environmental Assessment Office:
Coastal GasLink and LNG Canada assessments. CER and NRCan figures used in the model are cited
per cell in `Inputs/Canada_LNG_Asset_Register.xlsx` and `Inputs/Canada_LNG_Data_Inputs.xlsx`.

**Peer-reviewed literature**
- Davis, S.J., Peters, G.P. and Caldeira, K. (2011). The supply chain of CO2 emissions. *PNAS*
  108(45), 18554–18559. https://doi.org/10.1073/pnas.1107409108
- Andrew, R.M., Davis, S.J. and Peters, G.P. (2013). Climate policy and dependence on traded
  carbon. *Environmental Research Letters* 8, 034011. https://doi.org/10.1088/1748-9326/8/3/034011
- Steininger, K.W., Lininger, C., Meyer, L.H., Muñoz, P. and Schinko, T. (2016). Multiple carbon
  accounting to support just and effective climate policies. *Nature Climate Change* 6, 35–41.
  https://doi.org/10.1038/nclimate2867
Howarth (2024), *Energy Science and Engineering*. MacKay et al. (2021), *Scientific Reports*.
Johnson et al. (2023), *Communications Earth and Environment*. Balcombe et al. (2022),
*Environmental Science and Technology*. Di Lullo et al. (2020), *Journal of Natural Gas Science
and Engineering*.

**Industry**
International Gas Union, World LNG Report 2026. International Maritime Organization, Data
Collection System annual report. Intergovernmental Panel on Climate Change, AR6 and the 2006
Guidelines.

---

## Citation

Code, asset register and results:

> Stretch, F. (2026). *Canada LNG Expansion: Lifecycle Emissions Model*
> (version 2026.09.29). Zenodo. https://github.com/FStretch/canada_lng_carbon_bomb

The citable version is the Zenodo record created from the GitHub release. Use the DOI on
that record in the manuscript data-availability statement.

Manuscript this deposit accompanies:

> Stretch, F. (2026). *A Canadian Carbon Bomb? The potential lifetime emissions and
> economic damages of Canada's LNG export expansion.* Manuscript prepared for
> *Environmental Research Letters*.

---

## Contributing and corrections

Corrections are welcome, particularly on the asset register. If a capacity, status, licence term or
start year is wrong, please open an issue with the source that supports the correction. The
register records where each value came from, so a disagreement can usually be resolved by comparing
sources directly.
