# Model guide

This guide walks through `model/NREP_Tariff_LCOE_Model.xlsx` tab by tab. It explains what each calculation does and why it is built that way. If you can follow the Najran example in section 4 with a calculator, you can defend every number in the deck.

## 1. What the model answers

Saudi Arabia buys solar power at 1.1 to 1.5 US cents per kWh. IRENA puts the 2025 global average cost of utility solar at USD 44/MWh (4.4 cents), and its own Saudi average at USD 40/MWh. The model asks three things:

1. What would have to be true about cost, sunshine and financing for a bid like Najran (1.09682 cents, Masdar, October 2025) to earn a normal return?
2. How much of the gap to the global average comes from each of those?
3. If Oman's single buyer (Nama PWP) ran the same kind of tender, what tariff should it expect?

## 2. Tab order and flow

`Inputs` → `Tracker` and `Yield` → `Engine` (with `Engine_CF`) → `Sensitivity`, `Bridge`, `Oman` → `Impact`, `Value_chain`, `Changes`. `Capex` collects cost evidence. `PF_Model` is a levered view of one project. `Checks` ties the tabs together. `Cover` shows the headline numbers and the combined check.

Colour code: blue text is an input or reported figure, green text is a link from another tab, black text is a calculation. Yellow fill marks an assumption with no direct source. Green fill marks an output.

## 3. Inputs

Every input sits on one row of `Inputs` with a base value, a low and a high, a unit, a type and a source ID from `Sources`. The type tells you how much weight it can bear:

- Reported: copied from a source (for example, the Najran tariff, IRENA's USD 667/kWdc global installed cost, Damodaran's country risk premiums).
- Derived: calculated from reported figures (for example, the 72 to 78% gearing seen in two financial closes, or the yield calibration below).
- Assumed: a judgement, with a range. The debt margin, O&M cost and equity hurdle are assumed because no public source gives them for Saudi projects.

Every input has a workbook name equal to its ID, so formulas read `=CAPEX` rather than `=Inputs!$D$9`. The `_LO` and `_HI` names point at the range ends.

The derived block at the bottom of `Inputs` holds six formulas:

| Name | Formula | Najran case value |
|---|---|---|
| YIELD_CAL | Sakaka 2022 output (932 GWh) / 300 MWac / PVGIS yield at Sakaka | 1.2089 |
| KD_KSA | Swap rate + margin | 5.25% |
| WACC_KSA | 75% × 5.25% × (1 − 10%) + 25% × 9% | 5.79% |
| REAL_KSA | (1 + 5.79%) / (1 + 2.3%) − 1 | 3.42% |
| KD_OMN | KD_KSA + Oman default spread − Saudi default spread | 6.61% |
| KE_OMN | Equity hurdle + Oman country risk premium − Saudi premium | 11.07% |

### Why the yield calibration matters

PVGIS gives the output of 1 kWp of modules on a single-axis tracker at each site, after temperature, angle and spectral losses but before soiling, inverter, clipping and availability. A tariff is paid per kWh delivered per kWac of plant. Sakaka is a 300 MWac tracker plant whose measured 2022 output GASTAT publishes, so dividing that output by its PVGIS yield gives one factor that converts "PVGIS kWh per kWp" into "delivered kWh per kWac". The factor (1.209) bundles the plant's DC/AC ratio (about 1.35) and its net losses (about 10%). It is one plant from 2020, so the sensitivity tab runs it from 1.15 to 1.35.

## 4. The engine, column by column (Najran row)

Each row asks: at what capex would this tariff earn exactly the hurdle WACC after tax? Every step is a column so it can be audited.

Notation: T tariff (USD/kWh), E year-1 energy (kWh/kWac), O year-1 O&M (USD/kWac), r WACC, n PPA years, d degradation, g O&M escalation, τ tax rate, D depreciation years.

**Present value of energy (columns N, O).** Energy falls by d each year, so year t produces E × (1 − d)^(t−1). Discounted, the sum of that series has a closed form. With q = (1 − d) / (1 + r):

PV factor = q × (1 − q^n) / ((1 − q) × (1 − d))

Najran: q = 0.996 / 1.05794 = 0.94145, PV factor = 12.572. Revenue PV per kWac is T × E × 12.572 = 0.0109682 × 3,244 × 12.572 = USD 447.3.

**Present value of O&M (columns P, Q).** Same formula with growth g instead of decline. q = 1.02 / 1.05794 = 0.96414, PV factor = 15.780, so O&M PV = 8 × 15.780 = USD 126.2.

**Tax shield (columns R, S).** Straight-line depreciation of capex C over D years saves τ × C / D a year. Its PV is τ × C × annuity(r, D) / D. Column S holds τ × annuity / D = 0.10 × 13.038 / 25 = 0.0522, so the shield is 5.2% of capex.

**Implied capex (column T).** After-tax value equals capex when:

C = (1 − τ) × (T × E × PV_energy − O × PV_O&M) / (1 − shield)

Najran: 0.9 × (447.3 − 126.2) / (1 − 0.0522) = USD 305/kWac.

**Required tariff at base capex (column U).** The same equation solved for T with C = USD 520:

T = (C × (1 − shield) + (1 − τ) × O × PV_O&M) / ((1 − τ) × E × PV_energy) = 1.65 cents.

So at today's disclosed costs and a 5.8% hurdle, a Najran-quality site needs about 1.65 cents. Masdar bid 1.10.

**Disclosed capex and IRR (columns V to X).** Column V looks up the average disclosed Saudi cost for the tariff's vintage year on `Capex` (USD 510/kWac for 2025). Column X is the IRR of the cash flows on `Engine_CF` at that cost: 1.13% after tax, nominal.

**Real-equivalent tariff (columns Y to AA).** A flat 1.097 cents for 25 years is worth less each year in real terms. The constant real price with the same present value is 0.863 cents in year-0 money. That is the like-for-like figure to set against IRENA's real LCOE.

**Tariff at reference yield (column AB).** Tariff × E / average Saudi E. It removes the sunshine advantage so tariffs can be compared on cost and returns alone.

### Rebuild exercise

On a blank sheet, rebuild the Najran row from the inputs only: E, the two PV factors, the shield, implied capex and required tariff. You have it right when you hit 3,244 kWh, 12.572, 15.780, 0.0522, USD 305 and 1.652 cents. Then change CAPEX to 450 on `Inputs` and check that your required tariff moves the same way as `Engine!U` for Najran.

## 5. Engine_CF

One row of annual cash flows per tariff: year 0 is minus the disclosed capex for that vintage, and years 1 to n are EBITDA less tax on (EBITDA − depreciation). The IRR in column AP feeds `Engine!X`. Rows outside 2021 to 2025 have no disclosed cost and stay blank. The Checks tab rebuilds Najran's cash flows at the implied capex and confirms the IRR comes back to exactly the hurdle WACC, which proves the closed form and the cash flows agree.

## 6. PF_Model

Pick a project in C4. The model runs 35 annual columns per kWac.

1. CFADS is EBITDA less tax computed without interest.
2. Target debt service is CFADS / minimum DSCR during the tenor.
3. Debt capacity is the PV of that service at the cost of debt. Debt raised is the lower of capacity and 80% of capex.
4. Actual service is the target scaled down if gearing binds. The balance rolls forward, interest is charged on the opening balance, and cash tax uses profit after interest.
5. Equity cash flow is EBITDA less debt service less cash tax.

Sculpting on CFADS before the interest tax shield avoids a circular reference. The shield then goes to equity.

For Najran at USD 510/kWac, the DSCR test only supports 51% gearing, the project IRR is 1.1%, and equity IRR is −2.8%. A 1.1% unlevered return is below the 5.25% cost of debt, so leverage hurts. This is the result to explain in an interview: a commercial investor could not make this bid at disclosed costs.

## 7. Sensitivity

The top table recomputes Najran's required tariff with one input at a time at its low and high value. Swings, largest first: capex (1.47 to 1.99 cents), PPA length (1.52 at 35 years, 1.79 at 20), the yield factor (1.48 to 1.74), O&M, the equity hurdle, then the debt terms. The grid below shows required tariff for capex 350 to 650 against WACC 3% to 9%. Najran's 1.10 cents needs capex of USD 350/kWac at a 4.5% WACC, or USD 400 at about 3.3%. Both are well below the disclosed cost and the hurdle.

## 8. Bridge (IRENA global average to Najran)

Both ends are put on IRENA's basis: per kWdc, real terms, no tax, no degradation, LCOE = (capex × capital recovery factor + O&M) / energy.

- End A uses IRENA's 2025 global figures (USD 667/kWdc, 16.5% capacity factor, USD 12.2/kWdc O&M). The model solves for the real return that turns them into USD 44/MWh: 5.84%.
- End B uses Najran's inputs divided by the DC/AC ratio. The model solves for the real return that turns them into Najran's real-equivalent tariff: −0.34%.

Four things differ between A and B: capex, energy, O&M, and the return (with life). Changing them one after another gives a different bar for each depending on the order. The Shapley table runs all 16 combinations and averages each driver's effect over every order, so the bars add up and no ordering choice is hidden. Result, USD/MWh: 44.0, then energy −11.9, return −11.6, capex −8.7, O&M −3.3, giving 8.6 real. Adding back the flat nominal basis (+2.4) gives Najran's 11.0.

About 58% of the gap is sunshine and cost (things a buyer can find or tender for), and about a third is the return the sponsor accepted.

## 9. Oman

Same engine, two columns: a Najran-quality Saudi site at the Saudi hurdle (1.65 cents) and Ibri in Oman at Oman's hurdle (2.35 cents). Oman's financing adds the gap in sovereign default spread (Baa3 vs Aa3) to debt and the gap in country risk premium to equity. A 16-row Shapley table splits the 0.70-cent difference into resource (0.19), capex and scale (0.15), financing and tax (0.22) and contract tenor (0.14).

The cases table then stacks the levers Nama PWP controls: credit support that gives Saudi-grade financing (2.13), a GW-scale tender with a 25-year PPA (1.83), and a sponsor willing to take Najran's return (1.23). The avoided-cost rows put fuel for a gas plant at 2.9 cents at the regulated gas price and 5.5 cents at the LNG export netback.

Oman's real tariffs are not published, so these are estimates, not a back-test.

## 10. Impact

This tab answers why Saudi Arabia accepts the deal. Each MWh of solar lets an oil-fired plant burn less, and the crude it would have burned can be exported.

- Output: GW × average Saudi energy per kWac (3,092 kWh) gives 55.0 TWh a year for the 17.8 GW running in August 2026 (KAPSARC) and 106.5 TWh for the 34.45 GW of utility PV contracted.
- Oil displaced: TWh × heat rate (10 MMBtu/MWh) / 5.8 MMBtu per barrel × the share of solar output that displaces liquids (80%, assumed). For the fleet running now that is 75.9 million barrels a year, or 208 kb/d, 47% of 2025 direct crude burn (440 kb/d).
- Value: barrels × (Brent USD 86 less a USD 3 discount) = USD 6.3bn a year; USD 8.5bn at the World Bank's severe case of USD 115.
- Cost: TWh × the MW-weighted Saudi tariff (1.344 cents) = USD 0.74bn. Every dollar of tariff frees 8.5 dollars of oil.
- CO2: barrels × 0.43 t = 32.6 Mt a year.

The Oman block takes solar output from Nama PWP's 7-Year Statement (16.2 TWh in 2030, 23.3 TWh in 2031) and multiplies by the CCGT heat rate (7.0 MMBtu/MWh) to get gas saved. At the regulated gas price that is worth USD 676m in 2031; at the LNG netback, USD 1,273m. The rule of thumb behind the moves row: 0.1 cent on 1 TWh is USD 1m. Moves 1 to 3 (2.35 to 1.83 cents) on 23.3 TWh are worth about USD 122m a year.

## 11. Value_chain

Three blocks.

The cost stack splits Najran's required tariff at base capex (1.65 cents) into who earns it. Lifetime energy is the degraded output over 25 years (77,326 kWh/kWac) and lifetime O&M is USD 256/kWac. Capex (USD 520) is split by the bottom-up shares on `Capex`, so modules are 0.19 cents, trackers 0.12, inverters 0.05, balance of system and freight 0.14, EPC and fees 0.16, and O&M 0.33. Together these give the cost before any return: 1.00 cents, undiscounted. The return on capital (0.58) is the pre-tax tariff at the hurdle WACC less that cost, and tax (0.07) is the required tariff less the pre-tax one. The stack sums to the Engine's required tariff, which Check 17 tests. Najran's 1.10-cent bid leaves 0.09 cents for capital and tax.

The what-has-to-be-true grid recomputes the capex Najran's tariff supports for lives of 25, 30 and 35 years and WACCs from 3% to 7%. The best cell, 35 years at 3%, supports USD 476/kWac. Cells that reach the USD 510 disclosed cost turn green; none do. Lives beyond 25 years assume the PPA is extended at the same tariff.

The timeline lists dated milestones for Sakaka, Sudair, Al Sadawi, Najran and Ibri III, with a month index (`=VALUE(LEFT(D,4))*12+VALUE(MID(D,6,2))`) and months since each project's first milestone (`MINIFS`).

## 12. Changes and the sponsor league

`Changes` lays out 2018 to 2026 by year: the US 10-year Treasury average (FRED, a proxy for SAR swap rates under the peg), the average Saudi PV tariff awarded (`AVERAGEIFS` on `Engine`), the average disclosed cost by vintage and the number of awards. Rates rose from 1.45% in 2021 to 4.29% in 2025 while disclosed cost fell from USD 722 to 510/kWac. Using the Sensitivity slopes at Najran's site, the rate rise is worth about +0.25 cents and the cost fall about −0.55 cents. Below sit the delivery block (19.3 GW installed, 23.9 under development, 20.6 tendered, against a 130 GW target and a 74.2 GW forecast, so a 55.8 GW gap) and the storage block (battery rounds of 8 and 12 GWh, a 48 GWh target, firm solar-plus-storage at USD 54 to 82/MWh).

The sponsor league sits at the foot of `Tracker`. A lead sponsor column names each consortium's leader; `SUMIFS` adds MW and MW × tariff over analysed awards, which gives each sponsor's GW and MW-weighted tariff. ACWA Power has 25.35 GW at 1.311 cents, Masdar 5.4 GW at 1.331, against 1.344 for the whole programme.

## 13. Checks

Fourteen tests, all of which must show TRUE: engine and cash-flow parity, sensitivity base equals the engine, both bridges reconcile and sum, the PF debt is repaid and respects its covenants, every analysed tariff is present, the value-chain stack sums to the required tariff, and there are no error values on the engine, impact, value-chain and changes tabs. `Cover` shows the combined result.

## 14. Known limits

- One calibration plant for yield. Newer plants with higher DC/AC ratios and bifacial gain probably do better, which would raise implied capex.
- No merchant tail after the PPA. Extra years of output after year 25 would raise returns at low discount rates.
- Disclosed "total investment cost" may include development fees and EPC margins paid to sponsor affiliates. If so, the true build cost is lower than USD 510, and part of the return sits outside the project company.
- Tax is a single blended rate. The real mix of 20% corporate tax and 2.5% Zakat depends on each consortium's ownership.
- Debt margin, O&M and equity hurdle are assumptions with ranges, not observed terms.
- Oil displacement rests on one heat rate and an assumed liquids share of 80%. At 50% the barrels and their value fall by three-eighths. A saved barrel is valued at export prices; with Hormuz shut in 2026, that depends on room on the Red Sea route.
