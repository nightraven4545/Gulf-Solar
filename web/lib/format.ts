export const cents = (v: number, d = 2) => `${v.toFixed(d)}¢`;
export const pct = (v: number, d = 1) => `${(v * 100).toFixed(d)}%`;
export const usd = (v: number) => `USD ${Math.round(v).toLocaleString("en-US")}`;
export const fix = (v: number, d = 2) => Number(v.toFixed(d));
