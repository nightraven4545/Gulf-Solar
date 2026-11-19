"use client";

import { Plot } from "@/lib/plot";
import { useEffect, useRef } from "react";

export const C = {
  ink: "#1f2a36", ink2: "#47515c", muted: "#66707b", rule: "#d9dee3", grid: "#e9ecef", wash: "#f3f5f7",
  amber: "#c2780b", blue: "#3a5fa8", teal: "#1a9a8a", grey: "#a3acb5", sun0: "#fff4dc", sun1: "#f2c46e", sun2: "#c2780b",
};
const FONT = '"Archivo Variable", "Helvetica Neue", Arial, sans-serif';
export const style = { fontFamily: FONT, fontSize: "13px", color: C.ink2, background: "transparent", overflow: "visible" };
const track = (t: string, country = "KSA") => (country === "Oman" ? C.teal : t === "SPPC" ? C.amber : t === "PIF" ? C.blue : C.grey);

/** Draws a Plot figure into a div and redraws it when the div's width changes. */
export function Chart({ render, minHeight = 240 }: { render: (width: number) => Element; minHeight?: number }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    let last = 0;
    const draw = () => {
      const w = Math.round(el.clientWidth);
      if (!w || w === last) return;
      last = w;
      el.replaceChildren(render(w));
    };
    draw();
    const ro = new ResizeObserver(draw);
    ro.observe(el);
    return () => ro.disconnect();
  }, [render]);
  return <div className="chart" ref={ref} style={{ minHeight }} />;
}

const narrow = (w: number) => w < 640;

// Exhibit 2 -------------------------------------------------------------------------------------------
type Award = { id: string; project: string; country: string; track: string; date: string; t: number; mw: number; tariff: number };
export function HistoryChart({ points, irena, irenaKsa }: { points: Award[]; irena: number; irenaKsa: number }) {
  const show = new Set(["NAJRAN", "SHUAIBAH1", "SAKAKA"]);
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: narrow(w) ? 320 : 380, marginLeft: 40, marginRight: narrow(w) ? 12 : 150, style,
          x: { domain: [2017.6, 2026.4], label: null, tickFormat: "d", ticks: [2018, 2020, 2022, 2024, 2026] },
          y: { domain: [0, 4.8], label: "US¢/kWh", grid: true, ticks: 5, tickFormat: (d: number) => d.toFixed(1) },
          marks: [
            Plot.ruleY([irena, irenaKsa], { stroke: C.grey, strokeDasharray: "4,3" }),
            ...([[irena, "IRENA global", -7], [irenaKsa, "IRENA Saudi", 7]] as const).map(([v, name, dy]) =>
              Plot.text([v], { x: 2026.4, y: (d: number) => d, dx: 6, dy, textAnchor: "start", fill: C.ink2, text: (d: number) => `${name} ${d.toFixed(1)}¢`, filter: () => !narrow(w) }),
            ),
            Plot.dot(points, {
              x: "t", y: "tariff", r: 5.5, fill: (d: Award) => track(d.track, d.country), stroke: "#fff", strokeWidth: 1.5,
              channels: { Project: "project", Awarded: "date", "US¢/kWh": "tariff", MWac: "mw" },
              tip: { format: { x: false, y: false, fill: false, Project: true, Awarded: true, "US¢/kWh": (v: number) => v.toFixed(3), MWac: true } },
            }),
            Plot.text(points.filter((d) => show.has(d.id)), {
              x: "t", y: "tariff", text: (d: Award) => `${d.project} ${d.tariff.toFixed(2)}¢`, dx: 9, textAnchor: "start",
              fill: C.ink, stroke: "#fff", strokeWidth: 3,
            }),
          ],
        })
      }
    />
  );
}

// Exhibit 3 -------------------------------------------------------------------------------------------
type Bar = { fleet: string; series: string; value: number };
export function ImpactChart({ bars, series }: { bars: Bar[]; series: string[] }) {
  const colors = [C.grey, C.amber, C.sun1];
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: 330, marginLeft: 44, marginBottom: narrow(w) ? 48 : 36, style,
          fx: { label: null, padding: 0.18, domain: [...new Set(bars.map((b) => b.fleet))], axis: null },
          x: { axis: null, domain: series, padding: 0.12 },
          y: { label: "USD bn a year", grid: true, ticks: 5 },
          color: { domain: series, range: colors },
          marks: [
            Plot.barY(bars, { fx: "fleet", x: "series", y: "value", fill: "series", rx: 2, tip: { format: { fx: true, x: false, y: (v: number) => `$${v.toFixed(1)}bn`, fill: true } } }),
            Plot.text(bars, { fx: "fleet", x: "series", y: "value", text: (d: Bar) => `$${d.value.toFixed(1)}bn`, dy: -8, fill: C.ink }),
            Plot.ruleY([0], { stroke: C.ink }),
            Plot.axisFx({ lineWidth: narrow(w) ? 11 : 30, tickSize: 0, fill: C.ink, fontSize: 13, label: null }),
          ],
        })
      }
    />
  );
}

// Exhibit 4 -------------------------------------------------------------------------------------------
type Step = { name: string; x1: number; x2: number; kind: "part" | "total" | "hurdle" | "bid" };
export function ValueChainChart({ steps }: { steps: Step[] }) {
  const fill = { part: C.grey, total: C.ink, hurdle: C.sun1, bid: C.amber };
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: 40 * steps.length + 40, marginLeft: narrow(w) ? 150 : 270, marginRight: 40, style,
          x: { domain: [0, 1.8], label: null, grid: true, ticks: narrow(w) ? [0, 0.5, 1, 1.5] : [0, 0.2, 0.4, 0.6, 0.8, 1, 1.2, 1.4, 1.6, 1.8], tickFormat: (d: number) => `${d.toFixed(1)}¢` },
          y: { domain: steps.map((s) => s.name), label: null, axis: null },
          marks: [
            Plot.axisY({ tickSize: 0, lineWidth: narrow(w) ? 11 : 24, fill: C.ink, fontSize: 13 }),
            Plot.barX(steps, {
              y: "name", x1: "x1", x2: "x2", fill: (d: Step) => fill[d.kind], insetTop: 4, insetBottom: 4,
              channels: { Value: (d: Step) => d.x2 - d.x1 },
              tip: { format: { x1: false, x2: false, y: true, fill: false, Value: (v: number) => `${v.toFixed(3)}¢` } },
            }),
            Plot.text(steps, {
              y: "name", x: "x2", text: (d: Step) => (d.x2 - d.x1).toFixed(2), dx: 6, textAnchor: "start", fill: C.ink,
            }),
          ],
        })
      }
    />
  );
}

// Exhibit 5 -------------------------------------------------------------------------------------------
type Pt = { life: string; wacc: number; capex: number };
export function WhtbtChart({ points, disclosed, hurdle }: { points: Pt[]; disclosed: number; hurdle: number }) {
  const lives = [...new Set(points.map((p) => p.life))];
  const first = lives.map((l) => points.find((p) => p.life === l)!);
  const atHurdle = points.find((p) => p.life === lives[0] && Math.abs(p.wacc - hurdle) < 1e-9)!;
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: 390, marginLeft: 48, marginRight: 20, marginBottom: 48, style,
          x: { domain: [0.02, 0.074], label: "After-tax cost of capital (WACC)", labelAnchor: "center", labelOffset: 40, labelArrow: "none", tickFormat: (d: number) => `${(d * 100).toFixed(0)}%`, ticks: [0.03, 0.04, 0.05, 0.06, 0.07] },
          y: { domain: [250, 540], label: "Capex the bid supports, USD/kWac", grid: true },
          color: { domain: lives, range: [C.grey, C.blue, C.amber] },
          marks: [
            Plot.ruleY([disclosed], { stroke: C.ink, strokeDasharray: "5,4" }),
            Plot.text([disclosed], { x: 0.074, y: (d: number) => d, dy: -8, textAnchor: "end", fill: C.ink, text: (d: number) => `Disclosed 2025 cost: USD ${d}` }),
            Plot.line(points, { x: "wacc", y: "capex", stroke: "life", strokeWidth: 2 }),
            Plot.dot(points, {
              x: "wacc", y: "capex", fill: "life", r: 4.5, stroke: "#fff", strokeWidth: 1.5,
              tip: { format: { x: (v: number) => `${(v * 100).toFixed(1)}%`, y: (v: number) => `USD ${v.toFixed(0)}/kWac`, fill: true } },
            }),
            Plot.text(first, { x: "wacc", y: "capex", text: (d: Pt) => `${d.life}: ${d.capex.toFixed(0)}`, dx: -9, textAnchor: "end", fill: C.ink, filter: () => !narrow(w) }),
            Plot.text([atHurdle], { x: "wacc", y: "capex", text: (d: Pt) => `Hurdle ${(d.wacc * 100).toFixed(1)}%: ${d.capex.toFixed(0)}`, dy: 16, fill: C.ink, fontWeight: 600 }),
          ],
        })
      }
    />
  );
}

// Exhibit 6 -------------------------------------------------------------------------------------------
type Wf = { name: string; y1: number; y2: number; kind: "end" | "step" | "basis"; label: string };
export function BridgeChart({ steps }: { steps: Wf[] }) {
  const fill = { end: C.ink, step: C.amber, basis: C.grey };
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: narrow(w) ? 430 : 400, marginLeft: narrow(w) ? 64 : 44, marginBottom: narrow(w) ? 100 : 56, style,
          x: { domain: steps.map((s) => s.name), label: null, padding: 0.25 },
          y: { domain: [0, 48], label: "USD/MWh (real)", grid: true },
          marks: [
            narrow(w)
              ? Plot.axisX({ tickRotate: -45, tickSize: 0, fill: C.ink, textAnchor: "end" })
              : Plot.axisX({ lineWidth: 9, tickSize: 0, fill: C.ink }),
            Plot.barY(steps, {
              x: "name", y1: "y1", y2: "y2", fill: (d: Wf) => fill[d.kind], rx: 2,
              channels: { Effect: "label" }, tip: { format: { x: true, y1: false, y2: false, fill: false, Effect: true } },
            }),
            Plot.text(steps, { x: "name", y: (d: Wf) => Math.max(d.y1, d.y2), text: "label", dy: -8, fill: C.ink }),
            Plot.ruleY([0], { stroke: C.ink }),
          ],
        })
      }
    />
  );
}

// Exhibit 7 -------------------------------------------------------------------------------------------
type Range = { name: string; low: number; mid?: number; high: number; kind: "site" | "actual" };
export function RangeChart({ rows, domain }: { rows: Range[]; domain: [number, number] }) {
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: 56 * rows.length + 40, marginLeft: narrow(w) ? 120 : 200, marginRight: 20, style,
          x: { domain, label: null, grid: true, ticks: Array.from({ length: Math.round((domain[1] - domain[0]) * 10) + 1 }, (_, i) => domain[0] + i / 10).filter((_, i) => !narrow(w) || i % 2 === 0), tickFormat: (d: number) => `${d.toFixed(1)}¢` },
          y: { domain: rows.map((r) => r.name), label: null, axis: null },
          marks: [
            Plot.axisY({ tickSize: 0, lineWidth: narrow(w) ? 9 : 16, fill: C.ink, fontSize: 13 }),
            Plot.barX(rows, {
              y: "name", x1: "low", x2: "high", fill: (d: Range) => (d.kind === "site" ? C.sun1 : C.grey), insetTop: 10, insetBottom: 10, rx: 2,
              tip: { format: { y: true, x1: (v: number) => `${v.toFixed(3)}¢`, x2: (v: number) => `${v.toFixed(3)}¢`, fill: false } },
            }),
            Plot.tickX(rows.filter((r) => r.mid !== undefined), { y: "name", x: "mid", stroke: C.ink, strokeWidth: 2.5, insetTop: 6, insetBottom: 6 }),
            Plot.text(rows, {
              y: "name", x: "high", dx: 8, textAnchor: "start", fill: C.ink,
              text: (d: Range) => (d.mid !== undefined ? `${d.low.toFixed(2)}-${d.high.toFixed(2)}, mid ${d.mid.toFixed(2)}` : `${d.low.toFixed(2)}-${d.high.toFixed(2)} actual`),
              filter: () => !narrow(w),
            }),
          ],
        })
      }
    />
  );
}

// Exhibit 8 -------------------------------------------------------------------------------------------
type Case = { name: string; value: number; kind: "case" | "gas" };
export function OmanChart({ rows }: { rows: Case[] }) {
  return (
    <Chart
      render={(w) =>
        Plot.plot({
          width: w, height: 48 * rows.length + 40, marginLeft: narrow(w) ? 140 : 280, marginRight: 50, style,
          x: { domain: [0, 6], label: null, grid: true, ticks: [0, 1, 2, 3, 4, 5, 6], tickFormat: (d: number) => `${d}¢` },
          y: { domain: rows.map((r) => r.name), label: null, axis: null },
          marks: [
            Plot.axisY({ tickSize: 0, lineWidth: narrow(w) ? 10 : 26, fill: C.ink, fontSize: 13 }),
            Plot.barX(rows, { y: "name", x: "value", fill: (d: Case) => (d.kind === "case" ? C.teal : C.grey), insetTop: 6, insetBottom: 6, rx: 2, tip: { format: { y: true, x: (v: number) => `${v.toFixed(3)}¢`, fill: false } } }),
            Plot.text(rows, { y: "name", x: "value", text: (d: Case) => `${d.value.toFixed(2)}¢`, dx: 6, textAnchor: "start", fill: C.ink }),
            Plot.ruleX([0], { stroke: C.ink }),
          ],
        })
      }
    />
  );
}
