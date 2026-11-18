# Saudi solar tariffs: what drives a 1.1 ¢/kWh bid, and can Oman match it?

An outside-in case on public data. A Gulf utility (Oman's single buyer, Nama PWP) wants to know whether Saudi Arabia's record solar tariffs are replicable in its market: what drives a bid of about 1.1 US cents per kWh, and how sensitive is it to capital cost, capacity factor and financing?

## The answer

- **The record is real, but a tariff is not a cost.** Masdar won Najran (1.4 GW, October 2025) at 1.09682 ¢/kWh. Restated as a constant real price that is 0.86 ¢. IRENA's own 2025 average cost for Saudi solar is 4.0 ¢ and its global average is 4.4 ¢.
- **About 60% of the gap is sunshine and cost, about a third is the return sponsors accept.** At the 2025 disclosed cost of USD 510/kWac, Najran returns 1.1% a year before leverage, below a 5.25% cost of debt. Bridging IRENA's USD 44/MWh to Najran's USD 8.6 (real), energy yield explains USD 12, accepted return 12, capex 9 and O&M 3.
- **Oman can copy the design, not the returns, and still beat gas.** At commercial returns a 500 MW plant at Ibri needs about 2.35 ¢. Credit support, GW-scale tenders and 25-year PPAs bring that to about 1.83 ¢. Both sit below the 2.9 ¢ fuel cost of gas-fired power at Oman's regulated gas price, and well below the 5.5 ¢ value of that gas exported as LNG.

![Bridge from IRENA's global average to Najran](docs/waterfall.png)

## What is here

| Path | What it is |
|---|---|
| `model/NREP_Tariff_LCOE_Model.xlsx` | The model. 14 tabs, live formulas, no macros. Start on the Cover tab |
| `deck/NREP_Tariff_Replicability.pptx` and `.pdf` | 14 slides plus appendix; every number is read from the model |
| `onepager/NREP_onepager.pdf` | One-page summary |
| `docs/model-guide.md` | Tab-by-tab and formula-by-formula walkthrough, with a worked Najran example |
| `data/tracker.csv` | 37 projects: Saudi PV awards of 200 MW and above since 2018 from both buying tracks, plus UAE, Qatar and Oman comparators |
| `data/capex_evidence.csv` | Disclosed project costs from ACWA Power and Masdar filings |
| `data/assumptions.csv` | 47 inputs, each tagged Reported, Derived or Assumed, with a low and high value and a source |
| `data/sources.csv` | 39 sources with URLs and access dates |
| `data/yield.csv`, `data/raw/` | PVGIS site yields and the raw API responses they came from |
| `src/nrep/` | Python that fetches public data, builds the workbook and deck, and checks them |
| `tests/` | 21 tests, including a check that every workbook output matches an independent calculation |

## Method

1. **Tariffs.** Each award is taken from the buyer's announcement as reported by the trade press or the winning developer, in US cents and, where published, halalas. Tariffs that sources disagree on, or that were never published, stay in the tracker but out of the analysis.
2. **Sunshine.** PVGIS 5.3 gives the yield of a single-axis tracker at each site. One factor converts that to delivered AC energy per kWac. It is calibrated on Sakaka, a 300 MWac tracker plant whose 2022 output (932 GWh) GASTAT publishes.
3. **Cost.** Four views of capex: disclosed total investment per project (ACWA Power's FY2025 investor report, Masdar), IRENA's 2025 Saudi figure, a bottom-up build at October 2026 module prices, and the capex each tariff implies.
4. **Engine.** A closed-form after-tax levelised cost for every award gives the capex at which the tariff earns a 5.8% hurdle WACC, the project return at the disclosed cost, and the real-equivalent tariff. A single-asset project finance model adds sculpted debt and equity returns.
5. **Attribution.** A Shapley decomposition splits the gap between IRENA's global average and Najran across energy, capex, O&M and return. It averages each driver's effect over every order of change, so no ordering choice is hidden.
6. **Oman.** The same engine runs on Ibri's yield, Oman's disclosed Ibri III cost and financing adjusted for Oman's sovereign risk (Damodaran, January 2026). A second Shapley split separates what Oman inherits from what its tender design controls.

## Limits

- Yield rests on one calibration plant from 2020. Newer bifacial plants at higher DC/AC ratios probably produce more, which would raise the capex each tariff supports by about USD 40/kWac for every 10% of extra energy.
- Disclosed "total investment cost" may include development fees and EPC margins paid to sponsor affiliates. If so, the true build cost is lower and part of the sponsors' return sits outside the project company.
- No value is given to output after the 25-year PPA.
- Debt margin, O&M cost and equity hurdle are assumptions with ranges; no public source gives them for Saudi solar.
- Nama PWP does not publish its solar tariffs, so the Oman figures are estimates, not a back-test.
- Round 7 (3.1 GW) bids are in but unannounced as of 18 November 2026.

## Rebuild

Needs Python 3.12+ with `uv`, and LibreOffice for recalculation and PDF export.

```
make data     # pull PVGIS yields (and KAPSARC's project list) into data/
make model    # build the workbook and recalculate it
make deck     # build the deck and one-pager from the recalculated workbook
make check    # lint and tests
```

Code is MIT licensed. Data files are compiled from the public sources listed in `data/sources.csv`; cite those sources when reusing figures.
