"use client";

import { useMemo, useState } from "react";
import { breakevenGas, tariff, type Inputs } from "@/lib/lcoe";

export type Preset = Inputs & { key: string; label: string; expected: number };
type Site = { id: string; project: string; energy: number };
type Mark = { label: string; value: number; color: string };

function Field({ label, value, shown, min, max, step, onChange }: {
  label: string; value: number; shown: string; min: number; max: number; step: number; onChange: (v: number) => void;
}) {
  const id = label.replace(/\W+/g, "-").toLowerCase();
  return (
    <div className="field">
      <label htmlFor={id}>{label}</label>
      <output htmlFor={id}>{shown}</output>
      <input id={id} type="range" min={min} max={max} step={step} value={value} onChange={(e) => onChange(Number(e.target.value))} />
    </div>
  );
}

function Scale({ value, marks, max }: { value: number; marks: Mark[]; max: number }) {
  const x = (v: number) => `${(Math.min(v, max) / max) * 100}%`;
  return (
    <svg className="scale" width="100%" height="104" role="img" aria-label={`Required tariff ${value.toFixed(2)} cents against reference points`}>
      <rect x="0" y="34" width={x(value)} height="14" rx="2" fill="#1f2a36" />
      <line x1="0" x2="100%" y1="48.5" y2="48.5" stroke="#d9dee3" />
      {marks.map((m, i) => (
        <g key={m.label}>
          <line x1={x(m.value)} x2={x(m.value)} y1="24" y2="58" stroke={m.color} strokeWidth="2.5" />
          <text x={x(m.value)} y={i % 2 ? 76 : 18} textAnchor={m.value > max * 0.8 ? "end" : "middle"}>{`${m.label} ${m.value.toFixed(2)}¢`}</text>
        </g>
      ))}
      {[0, 1, 2, 3, 4, 5, 6].filter((t) => t <= max).map((t) => (
        <text key={t} x={x(t)} y="98" textAnchor={t === 0 ? "start" : t === max ? "end" : "middle"}>{`${t}¢`}</text>
      ))}
    </svg>
  );
}

export default function Calculator({ presets, sites, marks, heatRate }: { presets: Preset[]; sites: Site[]; marks: Mark[]; heatRate: number }) {
  const [active, setActive] = useState(presets[0].key);
  const [v, setV] = useState<Inputs>(presets[0]);
  const set = (k: keyof Inputs) => (x: number) => {
    setActive("");
    setV((o) => ({ ...o, [k]: x, ...(k === "years" ? { dep: x } : {}) }));
  };
  const t = useMemo(() => tariff(v), [v]);
  const site = sites.find((s) => Math.abs(s.energy - v.energy) < 0.5);

  return (
    <div>
      <div className="presets" role="group" aria-label="Load a case from the model">
        {presets.map((p) => (
          <button key={p.key} type="button" className="preset" aria-pressed={active === p.key}
            onClick={() => { setActive(p.key); setV(p); }}>
            {p.label}
          </button>
        ))}
      </div>
      <div className="calc">
        <div>
          <fieldset>
            <legend>Cost</legend>
            <Field label="Capex, USD per kWac" value={v.capex} shown={`${Math.round(v.capex)}`} min={250} max={900} step={5} onChange={set("capex")} />
            <Field label="O&M, USD per kWac a year" value={v.opex} shown={v.opex.toFixed(1)} min={4} max={16} step={0.5} onChange={set("opex")} />
          </fieldset>
          <fieldset>
            <legend>Performance</legend>
            <div className="field">
              <label htmlFor="site">Site (sets energy from PVGIS)</label>
              <span />
              <select id="site" value={site?.id ?? ""} onChange={(e) => {
                const s = sites.find((x) => x.id === e.target.value);
                if (s) set("energy")(s.energy);
              }}>
                <option value="">Custom</option>
                {sites.map((s) => <option key={s.id} value={s.id}>{s.project}</option>)}
              </select>
            </div>
            <Field label="Energy, kWh per kWac in year 1" value={v.energy} shown={`${Math.round(v.energy).toLocaleString("en-US")}`} min={2600} max={3500} step={5} onChange={set("energy")} />
          </fieldset>
          <fieldset>
            <legend>Finance</legend>
            <Field label="After-tax WACC" value={v.wacc} shown={`${(v.wacc * 100).toFixed(2)}%`} min={0.005} max={0.12} step={0.0005} onChange={set("wacc")} />
            <Field label="PPA length, years" value={v.years} shown={`${v.years}`} min={15} max={35} step={1} onChange={set("years")} />
            <Field label="Tax rate" value={v.tax} shown={`${(v.tax * 100).toFixed(1)}%`} min={0} max={0.25} step={0.005} onChange={set("tax")} />
          </fieldset>
        </div>
        <div aria-live="polite">
          <p className="result-sub" style={{ marginBottom: 0 }}>Required flat tariff over the PPA</p>
          <p className="result-n">{t.toFixed(3)}¢</p>
          <p className="result-sub">
            US cents per kWh. A gas plant burning fuel at {heatRate.toFixed(1)} MMBtu/MWh matches it at a gas price of USD {breakevenGas(t, heatRate).toFixed(2)}/MMBtu.
          </p>
          <Scale value={t} marks={marks} max={6} />
          <details>
            <summary>How it is calculated</summary>
            <p>
              Tariff = (capex × (1 − tax shield) + (1 − tax) × PV of O&amp;M) / ((1 − tax) × PV of energy), with every present value
              taken at the WACC over the PPA, output falling {(v.degr * 100).toFixed(1)}% a year and O&amp;M rising {(v.esc * 100).toFixed(0)}% a year.
              It is the formula in column U of the workbook&apos;s Engine tab, and each preset reproduces the workbook to nine decimal places.
            </p>
          </details>
        </div>
      </div>
    </div>
  );
}
