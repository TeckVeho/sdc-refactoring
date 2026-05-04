import { describe, expect, it } from "vitest";

import {
  kkmdRangeFromMonth,
  unoRangeFromYearMonth,
} from "./gamma-results-aggregate.js";

describe("gamma-results-aggregate", () => {
  it("unoRangeFromYearMonth matches Excel-style derivation for 2026-04", () => {
    const r = unoRangeFromYearMonth(2026, 4);
    expect(r.unoMin.length).toBe(10);
    expect(r.unoMax.length).toBe(10);
    expect(r.unoMin.endsWith("0000")).toBe(true);
    expect(r.unoMax.endsWith("9999")).toBe(true);
    expect(r.unoMin < r.unoMax).toBe(true);
  });

  it("kkmdRangeFromMonth pads month", () => {
    expect(kkmdRangeFromMonth(4)).toEqual({ kkmd1: "0400", kkmd2: "0432" });
    expect(kkmdRangeFromMonth(12)).toEqual({ kkmd1: "1200", kkmd2: "1232" });
  });
});
