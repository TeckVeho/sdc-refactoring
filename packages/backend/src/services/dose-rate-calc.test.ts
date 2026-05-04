import { describe, expect, it } from "vitest";

import {
  CABLE_CORRECTION_DEFAULT,
  CS_TO_CO_WATER_FACTOR,
  computeDoseRate,
  kappaTp,
  roundDown3Significant,
} from "./dose-rate-calc.js";
import { decayFactor } from "./radiation-decay.js";

describe("kappaTp", () => {
  it("22℃ 1013.3hPa で参考値 1 に近い", () => {
    const v = kappaTp(22, 1013.3);
    expect(v).toBe(1);
  });
});

describe("roundDown3Significant", () => {
  it("1.23456 → 1.234 相当", () => {
    expect(roundDown3Significant(1.23456)).toBeCloseTo(1.234, 6);
  });
});

describe("computeDoseRate (RATE)", () => {
  it("減衰なしで有限の積と表示値を返す", () => {
    const r = computeDoseRate({
      mode: "RATE",
      potentiometerId: 1,
      ionChamberId: 1,
      readValue: 10,
      temperatureC: 22,
      pressureHpa: 1013.3,
      tpEnabled: false,
      uncertaintyOrg: "none",
      cableOn: true,
    });
    expect(r.potentiometerConst).toBe(1.0013);
    expect(r.kappaTp).toBe(1);
    expect(r.cableFactor).toBe(CABLE_CORRECTION_DEFAULT);
    expect(r.uncertaintyFactor).toBe(1);
    expect(Number.isFinite(r.rawProduct)).toBe(true);
    expect(r.displayDoseRateGyPerH != null && Number.isFinite(r.displayDoseRateGyPerH)).toBe(true);
  });

  it("減衰係数を表示線量率に反映", () => {
    const r = computeDoseRate({
      mode: "RATE",
      potentiometerId: 1,
      ionChamberId: 1,
      readValue: 100,
      temperatureC: 22,
      pressureHpa: 1013.3,
      tpEnabled: false,
      uncertaintyOrg: "none",
      cableOn: false,
      decay: { halfLifeDays: 10, elapsedDays: 10 },
    });
    const f = decayFactor(10, 10);
    expect(r.decayFactor).toBe(f);
    if (r.displayDoseRateGyPerH != null) {
      expect(r.displayDoseRateAfterDecay).toBeCloseTo(r.displayDoseRateGyPerH * f, 6);
    }
  });
});

describe("constants", () => {
  it("Cs→Co 水換算が仕様通り 0.02965/0.02932", () => {
    expect(CS_TO_CO_WATER_FACTOR).toBeCloseTo(0.02965 / 0.02932, 8);
  });
});
