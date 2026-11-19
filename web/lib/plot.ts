import * as runtime from "@observablehq/plot";
import type * as PlotTypes from "plot-types";

/** Plot at runtime, typed with the declarations Plot ships in src/. */
export const Plot = runtime as unknown as typeof PlotTypes;
