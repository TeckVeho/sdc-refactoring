export function tempCorrection(absThickness: number, temperatureC: number): number {
  return (-3.2478 + 23.386 * absThickness) * (-0.006 * temperatureC + 1.0558);
}
