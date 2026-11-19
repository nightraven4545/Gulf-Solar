import Calculator, { type Preset } from "@/components/Calculator";
import { BridgeChart, HistoryChart, ImpactChart, OmanChart, RangeChart, ValueChainChart, WhtbtChart } from "@/components/charts";
import Exhibit, { type Source } from "@/components/Exhibit";
import MapChart from "@/components/MapChart";
import { cents, fix, pct } from "@/lib/format";
import data from "@/public/data/case.json";

const REPO = "https://github.com/nightraven4545/Gulf-Solar";
const SOURCES = data.sources as Record<string, { publisher: string; title: string; date: string; url: string }>;
const src = (...ids: string[]): Source[] => ids.map((id) => ({ id, url: SOURCES[id].url, title: `${SOURCES[id].publisher}: ${SOURCES[id].title}` }));

export default function Page() {
  const h = data.headline;
  const ms = data.mapStats;
  const wf = Object.fromEntries(data.bridge.map((b) => [b.name, b.value]));
  const start = wf["IRENA global average 2025"];
  const real = wf["Najran, real equivalent"];
  const end = wf["Najran awarded tariff"];
  const gap = start - real;
  const shareSun = -(wf["Energy yield"] + wf["Capex"]) / gap;
  const shareRet = -wf["Return accepted (and life)"] / gap;
  const op = data.impact.operating;
  const ct = data.impact.contracted;
  const o = data.oman;
  const r7 = data.round7;

  // Exhibit 4: Najran's required tariff, built up from what it costs
  let run = 0;
  const vc = data.valueChain;
  const parts = vc.parts.map((p) => ({ name: p.name, x1: run, x2: (run += p.value), kind: "part" as const }));
  const vcSteps = [
    ...parts,
    { name: "Cost before any return", x1: 0, x2: h.preFin, kind: "total" as const },
    { name: "Return on capital at the hurdle", x1: h.preFin, x2: h.preFin + vc.ret, kind: "hurdle" as const },
    { name: "Tax and Zakat", x1: h.preFin + vc.ret, x2: h.required, kind: "hurdle" as const },
    { name: "Required at the hurdle WACC", x1: 0, x2: h.required, kind: "total" as const },
    { name: "Najran awarded tariff", x1: 0, x2: h.tariff, kind: "bid" as const },
  ];

  // Exhibit 5: what has to be true
  const lives = data.whtbt.lives;
  const wt = lives.flatMap(([life, ...v]) => data.whtbt.waccs.map((w, j) => ({ life: `${life} years`, wacc: w, capex: v[j] })));
  const best = lives[lives.length - 1][1];
  const hIdx = data.whtbt.waccs.findIndex((w) => Math.abs(w - h.wacc) < 1e-9);
  const hole = h.capexDisclosed - lives[0][hIdx + 1];

  // Exhibit 6: Shapley bridge
  const steps = ["Energy yield", "Return accepted (and life)", "Capex", "O&M"].map((k) => ({ k, v: wf[k] })).sort((a, b) => a.v - b.v);
  let lvl = start;
  const bridge = [
    { name: "IRENA global 2025", y1: 0, y2: start, kind: "end" as const, label: start.toFixed(1) },
    ...steps.map(({ k, v }) => ({ name: k.replace(" (and life)", ""), y1: lvl + v, y2: (lvl += v) - v, kind: "step" as const, label: v.toFixed(1) })),
    { name: "Najran, real", y1: 0, y2: real, kind: "end" as const, label: real.toFixed(1) },
    { name: "Flat nominal basis", y1: real, y2: end, kind: "basis" as const, label: `+${(end - real).toFixed(1)}` },
    { name: "Najran tariff", y1: 0, y2: end, kind: "end" as const, label: end.toFixed(1) },
  ];

  const omanRows = [
    ...["1. Oman today: 500 MW, 20-year PPA", "2. Plus Saudi-grade financing", "3. Plus GW-scale tender, 25-year PPA", "4. Plus a Najran-style sponsor return"]
      .map((name, i) => ({ name, value: o.cases[i].value, kind: "case" as const })),
    { name: "Gas fuel cost, regulated price", value: h.gas, kind: "gas" as const },
    { name: "Gas fuel cost, LNG export value", value: h.lng, kind: "gas" as const },
  ];

  const impactSeries = ["Tariffs paid", `Oil freed at Brent USD ${data.impact.brent}`, `Oil freed at Brent USD ${data.impact.brentHi} (severe Hormuz case)`];
  const fleets: [string, typeof op][] = [[`Solar running now (${op.gw.toFixed(1)} GW)`, op], [`All contracted utility PV (${ct.gw.toFixed(1)} GW)`, ct]];
  const impactBars = fleets.flatMap(([fleet, f]) => [f.cost, f.value, f.value_hi].map((value, i) => ({ fleet, series: impactSeries[i], value })));

  return (
    <>
      <header className="bar">
        <div className="wrap">
          <strong>Gulf solar tariffs</strong>
          <nav aria-label="Downloads">
            <a href="/downloads/NREP_Tariff_Replicability.pdf">Deck (PDF)</a>
            <a href="/downloads/NREP_Tariff_LCOE_Model.xlsx">Model (Excel)</a>
            <a href="/downloads/NREP_onepager.pdf">One-pager</a>
            <a href={REPO}>Code and data</a>
          </nav>
        </div>
      </header>

      <main className="wrap">
        <section className="hero" aria-labelledby="question">
          <p className="award-kicker">Najran solar PV, 1.4 GW, SPPC Round 6, awarded October 2025</p>
          <p className="award">
            {h.tariff.toFixed(5)}
            <span className="unit">US cents per kWh, flat for 25 years: the lowest solar tariff Saudi Arabia has signed</span>
          </p>
          <h1 id="question">Can Oman buy solar power as cheaply as Saudi Arabia?</h1>
          <div className="prose">
            <p className="dek">
              A Gulf utility, Oman&apos;s single buyer Nama PWP, wants to know whether Saudi Arabia&apos;s record tariffs can be repeated in
              its market. This case rebuilds every Saudi award from public data to find what drives a 1.1-cent bid, and which parts of it travel.
            </p>
            <ul className="findings">
              <li>
                <b>The record is real, but it is not a cost.</b> At the 2025 disclosed cost of USD {h.capexDisclosed}/kWac, Najran&apos;s bid
                returns {pct(h.irr)} a year before leverage, below the {pct(h.kd, 2)} its lenders charge. In real terms it is {cents(h.real)}.
              </li>
              <li>
                <b>About {pct(shareSun, 0)} of the gap to the global average is sunshine and cost; {pct(shareRet, 0)} is the return sponsors accept.</b>{" "}
                Building and running Najran costs {cents(h.preFin)} before any return, so the bid leaves {cents(h.left)} for capital.
              </li>
              <li>
                <b>Oman can copy the design, not the returns, and still beat gas.</b> Credit support, GW-scale sites and 25-year PPAs take it from{" "}
                {cents(h.omanCommercial)} to {cents(h.omanMoves)}, below the {cents(h.gas)} fuel cost of its gas power and worth about USD{" "}
                {Math.round(o.movesValue)}m a year by 2031.
              </li>
            </ul>
            <p className="links">
              <a href="#calculator">Try the calculator</a>
              <a href="/downloads/NREP_Tariff_Replicability.pdf">Read the 28-slide deck</a>
              <a href="/downloads/NREP_Tariff_LCOE_Model.xlsx">Open the model</a>
            </p>
          </div>
        </section>

        <section className="part" aria-labelledby="part-a">
          <div className="part-head">
            <p className="part-letter">A</p>
            <h2 id="part-a">What Saudi Arabia built, and what it is worth</h2>
          </div>
          <div className="prose">
            <p>
              Saudi Arabia has contracted {h.contractedGw.toFixed(0)} GW of utility solar since 2018, through open SPPC rounds and deals negotiated
              with PIF-backed sponsors. Every plant sells to one buyer under a flat 25-year price. The sun is strong almost everywhere, and the
              cheapest bids come from the sunniest and largest sites.
            </p>
          </div>

          <Exhibit n={1} file="yield-and-projects.csv"
            title={`Sunshine is strongest inland and in the north-west; Najran's site beats ${pct(ms.najranBeats, 0)} of Saudi Arabia, while Oman's Ibri would beat only ${pct(ms.ibriBeatsSau, 0)}`}
            sub="Tracked PV yield on a 1° grid, with awarded projects. Hover a project for its tariff, sponsor and site yield."
            note={`PVGIS 5.3 (SARAH3) at ${ms.cells} grid points; projects and tariffs from the model's Tracker and Engine tabs; borders from Natural Earth`}
            sources={src("S35", "S14", "S66")}
            table={{ columns: ["Measure", "kWh/kWp a year"], rows: [
              ["Saudi grid, lowest cell", fix(ms.sauMin, 0)], ["Saudi grid, highest cell", fix(ms.sauMax, 0)], ["Najran site", fix(ms.najran, 0)],
              ["Samtah site", fix(ms.samtah, 0)], ["Ibri site (Oman)", fix(ms.ibri, 0)], ["Oman grid, highest cell (Dhofar)", fix(ms.omnMax, 0)]] }}>
            <MapChart />
          </Exhibit>

          <Exhibit n={2} file="saudi-solar-tariffs.csv"
            title={`Awarded tariffs have held between 1.0 and 1.8¢ since 2021, between a quarter and two-fifths of IRENA's ${h.irena.toFixed(1)}¢ global average`}
            sub="Each dot is an award, in US cents per kWh, flat for the PPA term"
            note="model Tracker and Engine tabs; IRENA Renewable Power Generation Costs in 2025 (real LCOE for plants commissioned in 2025)"
            sources={src("S14", "S21", "S43")}
            table={{ columns: ["Project", "Country", "Track", "Awarded", "MWac", "US¢/kWh"], rows: data.history.map((p) => [p.project, p.country, p.track, p.date, p.mw, fix(p.tariff, 5)]) }}>
            <p className="legend">
              <span><i style={{ background: "#c2780b" }} />SPPC tender</span>
              <span><i style={{ background: "#3a5fa8" }} />PIF bilateral</span>
              <span><i style={{ background: "#a3acb5" }} />UAE and Qatar tenders</span>
            </p>
            <HistoryChart points={data.history} irena={h.irena} irenaKsa={h.irenaKsa} />
          </Exhibit>

          <Exhibit n={3} file="oil-freed-by-solar.csv"
            title={`The ${op.gw.toFixed(1)} GW of solar now running frees about ${op.kbd.toFixed(0)} kb/d of crude, ${pct(op.share, 0)} of 2025 burn, worth USD ${op.value.toFixed(1)}bn a year against USD ${op.cost.toFixed(1)}bn of tariffs`}
            sub={`Every dollar of tariff frees about ${op.ratio.toFixed(1)} dollars of oil for export. Values assume each saved barrel can be exported; with Hormuz shut in 2026, that depends on room on the Red Sea route.`}
            note="model Impact tab; crude burn from JODI via S&P Global; installed capacity from KAPSARC; heat rate from EIA; CO2 factor from EPA; Brent from the World Bank"
            sources={src("S44", "S45", "S56", "S49", "S32")}
            table={{ columns: ["Fleet", "GW", "TWh a year", "kb/d of crude", "Share of 2025 burn", "Oil value USD bn", "Tariffs USD bn", "CO2 Mt"],
              rows: fleets.map(([name, f]) => [name, fix(f.gw, 1), fix(f.twh, 1), fix(f.kbd, 0), pct(f.share, 0), fix(f.value, 2), fix(f.cost, 2), fix(f.co2, 1)]) }}>
            <p className="legend">
              {impactSeries.map((s, i) => <span key={s}><i style={{ background: ["#a3acb5", "#c2780b", "#f2c46e"][i] }} />{s}</span>)}
            </p>
            <ImpactChart bars={impactBars} series={impactSeries} />
          </Exhibit>
        </section>

        <section className="part" aria-labelledby="part-b">
          <div className="part-head">
            <p className="part-letter">B</p>
            <h2 id="part-b">What the bids imply</h2>
          </div>
          <div className="prose">
            <p>
              At disclosed costs the bids barely pay back. The next four exhibits split the tariff by who earns it, test what would have to be
              true for it to work, attribute the gap to the global average, and use the same engine to forecast the round that has just closed.
            </p>
          </div>

          <Exhibit n={4} file="najran-cost-stack.csv"
            title={`Najran costs about ${cents(h.preFin)} to build and run before any return; its ${cents(h.tariff)} bid leaves ${cents(h.left)} for lenders, equity and tax, against ${cents(vc.ret + vc.tax)} at a commercial hurdle`}
            sub="US cents per kWh. Cost before return is lifetime capex and O&M over lifetime energy, undiscounted"
            note={`model Value_chain tab; Najran at base capex, split by the bottom-up shares on the Capex tab`}
            sources={src("S43", "S18", "S20", "S51", "S52", "S53", "S54")}
            table={{ columns: ["Slice", "US¢/kWh", "Who captures it"], rows: [
              ...vc.parts.map((p) => [p.name, fix(p.value, 3), ""]),
              ["Return on capital at the hurdle", fix(vc.ret, 3), vc.who[3][1]], ["Tax and Zakat", fix(vc.tax, 3), vc.who[4][1]],
              ["Required at the hurdle", fix(h.required, 3), ""], ["Najran awarded tariff", fix(h.tariff, 5), ""]] }}>
            <ValueChainChart steps={vcSteps} />
            <div className="table-wrap" style={{ marginTop: 16 }}>
              <table>
                <thead><tr><th>Slice</th><th>Who captures it</th></tr></thead>
                <tbody>{vc.who.map(([k, v]) => <tr key={k}><td>{k}</td><td>{v}</td></tr>)}</tbody>
              </table>
            </div>
          </Exhibit>

          <Exhibit n={5} file="what-has-to-be-true.csv"
            title={`No mix of plant life and return in reach gets Najran to its disclosed cost: even ${lives[lives.length - 1][0]} years at a 3% WACC supports only USD ${best.toFixed(0)}/kWac, against USD ${h.capexDisclosed}`}
            sub={`Capex at which Najran's tariff earns each WACC, after tax. At the ${pct(h.wacc)} hurdle over 25 years the hole is USD ${hole.toFixed(0)}/kWac, about USD ${((hole * data.whtbt.mw) / 1000).toFixed(0)}m on Najran's ${(data.whtbt.mw / 1000).toFixed(1)} GW`}
            note="model Value_chain tab; lives beyond 25 years assume the PPA is extended at the same tariff"
            sources={src("S43")}
            table={{ columns: ["Life (years)", ...data.whtbt.waccs.map((w) => `WACC ${pct(w)}`)], rows: lives.map(([l, ...v]) => [l, ...v.map((x) => fix(x, 0))]) }}>
            <p className="legend">
              {lives.map(([l], i) => <span key={l}><i style={{ background: ["#a3acb5", "#3a5fa8", "#c2780b"][i] }} />{l}-year life</span>)}
              <span>Dashed line: disclosed 2025 cost</span>
            </p>
            <WhtbtChart points={wt} disclosed={h.capexDisclosed} hurdle={h.wacc} />
          </Exhibit>

          <Exhibit n={6} file="irena-to-najran-bridge.csv"
            title={`Of the USD ${gap.toFixed(0)}/MWh gap from IRENA's global average to Najran, energy yield explains ${(-wf["Energy yield"]).toFixed(0)}, accepted return ${(-wf["Return accepted (and life)"]).toFixed(0)}, capex ${(-wf["Capex"]).toFixed(0)} and O&M ${(-wf["O&M"]).toFixed(0)}`}
            sub="USD/MWh on IRENA's basis (real, per kWdc, before tax). Each bar is the driver's average effect over all 24 orders of change (Shapley)"
            note="model Bridge tab; IRENA Renewable Power Generation Costs in 2025"
            sources={src("S21", "S14")}
            table={{ columns: ["Step", "USD/MWh"], rows: data.bridge.map((b) => [b.name, fix(b.value, 2)]) }}>
            <BridgeChart steps={bridge} />
          </Exhibit>

          <Exhibit n={7} file="round-7-forecast.csv" badge="Forecast made 19 November 2026, before any Round 7 award"
            title={`If Round 7 bidders accept Round 6 returns, the four sites should clear at ${cents(Math.min(...r7.sites.map((s) => s.low)))} to ${cents(Math.max(...r7.sites.map((s) => s.high)))}, most likely near ${cents(r7.midAvg)}`}
            sub={`Bars run from Najran's return (${pct(r7.rates.low)}) to the highest Round 6 return (${pct(r7.rates.high)}); the tick is the MW-weighted Round 6 return (${pct(r7.rates.mid)}). Same engine, 2025 disclosed cost, each site's PVGIS yield`}
            note="model Forecast tab; bids closed on 13 September 2026"
            sources={src("S66", "S67", "S35")}
            table={{ columns: ["Site", "Region", "MWac", "kWh/kWac", "Low ¢", "Mid ¢", "High ¢"],
              rows: r7.sites.map((s) => [s.project, s.region, s.mw, fix(s.energy, 0), fix(s.low, 3), fix(s.mid, 3), fix(s.high, 3)]) }}>
            <RangeChart domain={[0.9, 1.6]} rows={[
              ...r7.sites.map((s) => ({ name: `${s.project} (${(s.mw / 1000).toFixed(1)} GW)`, low: s.low, mid: s.mid, high: s.high, kind: "site" as const })),
              { name: "Round 6 awards, actual", low: r7.r6[0], high: r7.r6[1], kind: "actual" as const },
            ]} />
          </Exhibit>
        </section>

        <section className="part" id="calculator" aria-labelledby="part-try">
          <div className="part-head">
            <h2 id="part-try">Try it: what tariff does a plant need?</h2>
          </div>
          <div className="prose">
            <p>
              Move the inputs and the required tariff updates with the formula from the model&apos;s Engine tab. The buttons load the model&apos;s
              own cases; each reproduces the workbook to nine decimal places, which a test in the repository checks.
            </p>
          </div>
          <Calculator
            presets={data.presets as Preset[]}
            sites={data.sites}
            heatRate={o.heatRate}
            marks={[
              { label: "Najran bid", value: h.tariff, color: "#c2780b" },
              { label: "Oman gas fuel", value: h.gas, color: "#66707b" },
              { label: "IRENA global", value: h.irena, color: "#a3acb5" },
              { label: "LNG value", value: h.lng, color: "#a3acb5" },
            ]}
          />
        </section>

        <section className="part" aria-labelledby="part-c">
          <div className="part-head">
            <p className="part-letter">C</p>
            <h2 id="part-c">What Oman can copy</h2>
          </div>
          <div className="prose">
            <p>
              Oman has bought 2 GW of solar so far, in 500 MW lots on 15 to 20-year PPAs, and its buyer carries a Baa3 sovereign rating. Three tender
              choices close most of the gap to a Saudi-quality site. The fourth lever, bidder appetite, is the part Nama PWP cannot design. Its gas price
              is not published, but solar wins as long as gas costs more than USD {o.breakeven[0].toFixed(1)}/MMBtu today, or {o.breakeven[1].toFixed(1)} with
              the moves, against an assumed {o.gasPrice.toFixed(2)}.
            </p>
          </div>

          <Exhibit n={8} file="oman-cases.csv"
            title={`With Saudi-style tender design Oman could reach about ${cents(o.cases[2].value, 1)}, below the ${cents(h.gas, 1)} fuel cost of gas power at regulated prices and ${cents(h.lng, 1)} at LNG export value`}
            sub="US cents per kWh. Cases add one change at a time; gas bars are fuel cost only, before any capacity value"
            note="model Oman tab; gas price escalated from the 2015 regulated price; LNG at a Brent-linked netback; Oman risk from Damodaran"
            sources={src("S22", "S25", "S31", "S32")}
            table={{ columns: ["Case", "US¢/kWh"], rows: omanRows.map((r) => [r.name, fix(r.value, 3)]) }}>
            <OmanChart rows={omanRows} />
          </Exhibit>

          <Exhibit n={9} file="four-moves.csv" toggle={false}
            title={`Four moves for Nama PWP, in order of what it controls; the first three are worth about USD ${Math.round(o.movesValue)}m a year on the solar in its 2031 plan`}
            sub={`On ${o.twh2031.toFixed(0)} TWh of solar in 2031, every 0.1¢ off the tariff is worth USD ${o.per01.toFixed(0)}m a year. The gas that solar saves is worth USD ${Math.round(o.gasValue)}m at the regulated price and USD ${(o.lngValue / 1000).toFixed(2)}bn as LNG`}
            note="model Oman and Impact tabs; Nama PWP 7-Year Statement 2025-2031; Oman's record from the Tracker tab"
            sources={src("S55", "S60", "S62", "S64", "S25", "S22")}
            table={{ columns: ["Move", "What it changes", "Effect on tariff (¢)", "USD m a year, 2031", "Saudi precedent", "Oman so far"],
              rows: o.moves.map((m, i) => [`${i + 1}. ${m.move}`, m.what, fix(-m.cents, 2), fix(m.usdm, 0), m.saudi, m.oman]) }}>
            <div className="table-wrap">
              <table className="wide">
                <thead>
                  <tr><th>Move</th><th>What it changes</th><th className="n">Effect on tariff</th><th className="n">USD m a year, 2031</th><th>Saudi precedent</th><th>Oman so far</th></tr>
                </thead>
                <tbody>
                  {o.moves.map((m, i) => (
                    <tr key={m.move}>
                      <td>{i + 1}. {m.move}</td><td>{m.what}</td>
                      <td className="n">{i === 3 ? "up to " : ""}-{m.cents.toFixed(2)}¢</td>
                      <td className="n">{i === 3 ? "up to " : ""}{Math.round(m.usdm)}</td>
                      <td>{m.saudi}</td><td>{m.oman}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Exhibit>
        </section>

        <section className="part method" aria-labelledby="method">
          <div className="prose">
            <h2 id="method">Method, limits and data</h2>
            <p>
              Tariffs come from buyer announcements as reported by the trade press or the winning developer. Yields come from PVGIS 5.3 for a
              single-axis tracker, calibrated to Sakaka&apos;s measured 2022 output. Capex is triangulated from ACWA Power and Masdar disclosures,
              a bottom-up build and IRENA. The engine is a closed-form after-tax levelised cost; attribution uses Shapley values; Oman&apos;s
              financing adds Damodaran&apos;s sovereign spreads. The workbook has 15 integrity checks and the repository tests every output
              against an independent Python calculation.
            </p>
            <ul>
              <li>Yield rests on one calibration plant from 2020; newer bifacial plants probably produce more.</li>
              <li>Disclosed &ldquo;total investment&rdquo; may include fees paid to sponsor affiliates, so the true build cost may be lower.</li>
              <li>No value is given to output after the PPA, except in Exhibit 5.</li>
              <li>Debt margin, O&amp;M and the equity hurdle are assumptions with ranges; no public source gives them for Saudi solar.</li>
              <li>Oil displacement rests on one heat rate and an assumed 80% share of solar output displacing liquids.</li>
              <li>Nama PWP does not publish tariffs, so the Oman figures are estimates, not a back-test.</li>
            </ul>
            <p className="links">
              <a href="/downloads/NREP_Tariff_LCOE_Model.xlsx">Model (Excel, 18 tabs)</a>
              <a href="/downloads/NREP_Tariff_Replicability.pdf">Deck (PDF, 28 slides)</a>
              <a href="/downloads/NREP_onepager.pdf">One-pager (PDF)</a>
              <a href={`${REPO}/blob/main/docs/model-guide.md`}>Model guide</a>
              <a href={REPO}>Code and data on GitHub</a>
            </p>
            <details>
              <summary>All {Object.keys(SOURCES).length} sources</summary>
              <ul>
                {Object.entries(SOURCES).map(([id, s]) => (
                  <li key={id}>
                    {id}. {s.publisher}, <a href={s.url} target="_blank" rel="noopener noreferrer">{s.title}</a> ({s.date})
                  </li>
                ))}
              </ul>
            </details>
          </div>
        </section>
      </main>

      <footer className="foot">
        <div className="wrap">
          Outside-in analysis on public data, as of {new Date(data.asOf).toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" })}. Not affiliated with or endorsed by any buyer,
          sponsor or firm named. Code MIT licensed; figures belong to the sources cited. Built by Shreshtha.
        </div>
      </footer>
    </>
  );
}
