import { describe, expect, it } from "vitest";

import { tempCorrection } from "./temp-correction.js";

describe("tempCorrection", () => {
  it("calculates correction for absThickness=1 and 25C", () => {
    expect(tempCorrection(1, 25)).toBeCloseTo(18.24118156, 8);
  });

  it("calculates correction for absThickness=0.5 and 40C", () => {
    expect(tempCorrection(0.5, 40)).toBeCloseTo(6.88959416, 8);
  });

  it("returns negative value when absThickness is zero", () => {
    expect(tempCorrection(0, 20)).toBeCloseTo(-3.03929124, 8);
  });
});
