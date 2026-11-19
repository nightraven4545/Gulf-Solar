// Port of src/nrep/lcoe.py. The same closed form as the workbook's Engine tab, column U.

export type Inputs = {
  capex: number; // USD per kWac
  energy: number; // kWh per kWac in year 1
  opex: number; // USD per kWac per year
  degr: number; // output decline per year
  esc: number; // O&M escalation per year
  tax: number;
  years: number; // PPA term
  dep: number; // depreciation years
  wacc: number; // after-tax nominal
};

/** PV of a stream that starts at 1 in year 1 and grows at `growth`. */
export function pvFactor(rate: number, years: number, growth = 0): number {
  const q = (1 + growth) / (1 + rate);
  if (Math.abs(1 - q) < 1e-12) return years / (1 + rate);
  return (q * (1 - q ** years)) / ((1 - q) * (1 + growth));
}

/** Flat tariff in US cents per kWh that returns `wacc` after tax. */
export function tariff(i: Inputs): number {
  const shield = (i.tax * pvFactor(i.wacc, i.dep)) / i.dep;
  const pvEnergy = i.energy * pvFactor(i.wacc, i.years, -i.degr);
  const pvOpex = i.opex * pvFactor(i.wacc, i.years, i.esc);
  return (100 * (i.capex * (1 - shield) + (1 - i.tax) * pvOpex)) / ((1 - i.tax) * pvEnergy);
}

/** Gas price (USD/MMBtu) at which a gas plant's fuel alone costs the same per kWh. */
export const breakevenGas = (cents: number, heatRate: number) => (cents * 10) / heatRate;
