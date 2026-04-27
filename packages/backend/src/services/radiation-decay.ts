export function decayFactor(elapsedDays: number, halfLifeDays: number): number {
  if (halfLifeDays <= 0) {
    throw new RangeError("halfLifeDays must be greater than 0");
  }

  return 0.5 ** (elapsedDays / halfLifeDays);
}
