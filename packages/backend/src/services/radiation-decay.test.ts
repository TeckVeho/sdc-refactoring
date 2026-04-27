import { describe, expect, it } from "vitest";

import { decayFactor } from "./radiation-decay.js";

describe("decayFactor", () => {
  it("returns 1 when elapsed days are zero", () => {
    expect(decayFactor(0, 1921)).toBe(1);
  });

  it("returns 0.5 after one half-life", () => {
    expect(decayFactor(1921, 1921)).toBe(0.5);
  });

  it("throws a RangeError when half-life is not positive", () => {
    expect(() => decayFactor(10, 0)).toThrowError(RangeError);
  });
});
