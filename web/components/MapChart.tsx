"use client";

import { Plot } from "@/lib/plot";
import { useEffect, useState } from "react";
import { C, Chart, style } from "./charts";

type Feature = { type: "Feature"; properties: Record<string, unknown>; geometry: { type: string; coordinates: unknown } };
type Project = {
  id: string; project: string; country: string; track: string; status: string; round: string;
  mw: number; lat: number; lon: number; sponsor: string; tariff: number | null; yield: number | null;
};
type MapData = {
  borders: { type: "FeatureCollection"; features: Feature[] };
  grid: { lon: number; lat: number; y: number }[];
  projects: Project[];
};

const LABEL: Record<string, string> = {
  NAJRAN: "Najran", SHUAIBAH1: "Shuaibah 1", SAKAKA: "Sakaka", SADAWI: "Al Sadawi", SAMTAH: "Samtah", IBRI3: "Ibri III",
};
const fillOf = (p: Project) => (p.country === "Oman" ? C.teal : p.track === "PIF" ? C.blue : C.amber);

// Each half-degree display point as a square cell (the 1° PVGIS samples, shown by nearest neighbour).
// Rings run clockwise, as d3-geo expects; anticlockwise would mean "the globe minus this square".
const cell = (lon: number, lat: number, h: number, y: number): Feature => ({
  type: "Feature",
  properties: { y },
  geometry: { type: "Polygon", coordinates: [[[lon - h, lat - h], [lon - h, lat + h], [lon + h, lat + h], [lon + h, lat - h], [lon - h, lat - h]]] },
});

export default function MapChart() {
  const [data, setData] = useState<MapData | null>(null);
  useEffect(() => {
    fetch("/data/map.json").then((r) => r.json()).then(setData);
  }, []);
  if (!data) return <div className="chart" style={{ minHeight: 420 }} aria-busy="true" />;

  const land = { type: "FeatureCollection" as const, features: data.borders.features.filter((f) => ["SAU", "OMN"].includes(String(f.properties.iso))) };
  const cells = data.grid.map((g) => cell(g.lon, g.lat, 0.26, g.y));
  const awarded = data.projects.filter((p) => p.status !== "bidding").sort((a, b) => b.mw - a.mw);
  const bidding = data.projects.filter((p) => p.status === "bidding");
  const tip = {
    format: { x: false, y: false, r: false, fill: false, stroke: false, Project: true, MWac: true, "Lead sponsor": true, Tariff: true, "PVGIS yield": true },
  };
  const channels = {
    Project: "project",
    MWac: "mw",
    "Lead sponsor": "sponsor",
    Tariff: (p: Project) => (p.tariff ? `${p.tariff.toFixed(3)}¢/kWh` : p.status === "bidding" ? "Round 7: bids in, no award yet" : "not published"),
    "PVGIS yield": (p: Project) => (p.yield ? `${Math.round(p.yield).toLocaleString("en-US")} kWh/kWp` : "n/a"),
  };
  const color = { type: "linear" as const, domain: [2100, 2450, 2800], range: [C.sun0, C.sun1, C.sun2], clamp: true };

  return (
    <>
      <p className="legend">
        <span><i style={{ background: C.amber }} />SPPC tender</span>
        <span><i style={{ background: C.blue }} />PIF bilateral</span>
        <span><i style={{ background: C.teal }} />Oman (Nama PWP)</span>
        <span><i className="ring" />Round 7 site, bids in</span>
        <span>Circle area shows capacity</span>
      </p>
      <Chart
        minHeight={420}
        render={(w) => {
          const fig = document.createElement("div");
          fig.append(
            Plot.plot({
              width: w, style, projection: { type: "mercator", domain: land, inset: 8 },
              color, r: { domain: [0, 2600], range: [0, w < 640 ? 11 : 17] },
              marks: [
                Plot.geo(data.borders, { fill: C.wash, stroke: C.rule, strokeWidth: 0.6, clip: "frame" }),
                Plot.geo(cells, { fill: (f: Feature) => f.properties.y as number, clip: land }),
                Plot.geo(land, { stroke: C.ink, strokeWidth: 0.8 }),
                Plot.dot(awarded, { x: "lon", y: "lat", r: "mw", fill: fillOf, fillOpacity: 0.9, stroke: "#fff", strokeWidth: 1.2, channels, tip }),
                Plot.dot(bidding, { x: "lon", y: "lat", r: "mw", stroke: C.amber, strokeWidth: 2, channels, tip }),
                Plot.text(awarded.filter((p) => LABEL[p.id]), {
                  x: "lon", y: "lat", dx: 12, textAnchor: "start", fill: C.ink, stroke: "#fff", strokeWidth: 3, fontSize: 12,
                  text: (p: Project) => `${LABEL[p.id]}${p.tariff ? ` ${p.tariff.toFixed(2)}¢` : ""}`,
                }),
                Plot.text([{ lon: 35.2, lat: 25.9, t: "Red Sea" }, { lon: 50.6, lat: 27.6, t: "Arabian Gulf" }], {
                  x: "lon", y: "lat", text: "t", fill: C.muted, fontStyle: "italic", fontSize: 12,
                }),
              ],
            }),
            Plot.legend({ color: { ...color, label: "Tracked PV yield, kWh/kWp per year (PVGIS 5.3)" }, width: Math.min(360, w), style, ticks: 4 }),
          );
          return fig;
        }}
      />
    </>
  );
}
