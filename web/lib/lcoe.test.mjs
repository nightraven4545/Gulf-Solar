// Run with `npm test` (Node's built-in runner; Node 22.18+ strips the TypeScript types in lcoe.ts).
import assert from "node:assert/strict";
import { test } from "node:test";
import data from "../public/data/case.json" with { type: "json" };
import { tariff } from "./lcoe.ts";

for (const p of data.presets) {
  test(`calculator matches the workbook: ${p.label}`, () => {
    assert.ok(Math.abs(tariff(p) - p.expected) < 1e-9, `${tariff(p)} vs ${p.expected}`);
  });
}

test("calculator reproduces every Round 7 forecast cell", () => {
  const base = data.presets[1]; // Najran at its bid return: 2025 disclosed cost, Saudi terms
  for (const s of data.round7.sites) {
    for (const k of ["low", "mid", "high"]) {
      const got = tariff({ ...base, energy: s.energy, wacc: data.round7.rates[k] });
      assert.ok(Math.abs(got - s[k]) < 1e-9, `${s.id} ${k}`);
    }
  }
});
