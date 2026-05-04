import { describe, expect, it } from "vitest";

import { computeRic3AdditionalPasses } from "./ric3-dose-shortage.js";

describe("computeRic3AdditionalPasses", () => {
  it("standard は単純 ROUNDUP（切り上げ）", () => {
    expect(computeRic3AdditionalPasses("simple_roundup", 11, 10, 2)).toBe(3);
    expect(computeRic3AdditionalPasses("simple_roundup", 12, 10, 2)).toBe(3);
    expect(computeRic3AdditionalPasses("simple_roundup", 10, 10, 2)).toBe(2);
  });

  it("sumiden は余裕 < 2 のとき +1", () => {
    const rounding = "sumiden_margin_lt2" as const;
    expect(computeRic3AdditionalPasses(rounding, 14, 10, 2)).toBe(4);
    expect(computeRic3AdditionalPasses(rounding, 12, 10, 2)).toBe(3);
  });

  it("入力不備時は null", () => {
    expect(computeRic3AdditionalPasses("simple_roundup", 10, 0, 2)).toBeNull();
    expect(computeRic3AdditionalPasses("simple_roundup", 10, 10, 0)).toBeNull();
  });
});
