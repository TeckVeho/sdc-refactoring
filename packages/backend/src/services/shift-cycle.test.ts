import { describe, expect, it } from "vitest";

import { shiftCycleFromYearMonth } from "./shift-cycle.js";

describe("shift-cycle", () => {
  it("11日〜翌月10日の日数を返す（2026-04）", () => {
    const c = shiftCycleFromYearMonth(2026, 4);
    expect(c.startIso).toBe("2026-04-11");
    expect(c.endIso).toBe("2026-05-10");
    expect(c.dates[0]).toBe("2026-04-11");
    expect(c.dates[c.dates.length - 1]).toBe("2026-05-10");
    expect(c.dates.length).toBeGreaterThanOrEqual(28);
  });
});
