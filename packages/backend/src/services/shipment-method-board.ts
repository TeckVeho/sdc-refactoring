import type { PrismaClient } from "@prisma/client";

import {
  parseShipmentJson,
  syouhoMaxNumeric,
} from "./shipment-method-registry.js";

const LIST_CAP = 1000;

export type ShipmentBoardRow = {
  kaisyacd: string;
  kairname: string;
  coname: string;
  hikitoriCurrent: string;
  hikitoriDb: string;
  housyubeCurrent: string;
  housyubeDb: string;
  reportFlagMax: number;
  needsReport: boolean;
};

export async function fetchShipmentMethodBoard(prisma: PrismaClient): Promise<{
  rows: ShipmentBoardRow[];
  truncated: boolean;
}> {
  const [tokus, flagRows, exRows] = await Promise.all([
    prisma.tokumst.findMany({
      where: { kaisyacd: { lt: "2000" } },
      orderBy: { kaisyacd: "asc" },
      take: LIST_CAP,
      select: { kaisyacd: true, kairname: true, coname: true },
    }),
    prisma.sehmst.groupBy({
      by: ["kaisyacd"],
      _max: { syouho: true },
    }),
    prisma.exSeihinj.findMany({
      where: { rowkey: { startsWith: "shipment:" } },
      select: { rowkey: true, json: true },
    }),
  ]);

  const flagMap = new Map<string, number>();
  for (const r of flagRows) {
    flagMap.set(r.kaisyacd, syouhoMaxNumeric(r._max.syouho ?? null));
  }

  const exMap = new Map<string, { hikitori: string; housyube: string }>();
  for (const r of exRows) {
    const cd = r.rowkey.startsWith("shipment:") ? r.rowkey.slice("shipment:".length) : "";
    if (cd) exMap.set(cd, parseShipmentJson(r.json));
  }

  const rows: ShipmentBoardRow[] = tokus.map((t) => {
    const dbEx = exMap.get(t.kaisyacd) ?? { hikitori: "", housyube: "" };
    const flag = flagMap.get(t.kaisyacd) ?? 0;
    const needsReport = flag > 0;
    let displayHousyube = dbEx.housyube;
    let displayHikitori = dbEx.hikitori;
    if (!needsReport) {
      displayHousyube = "不要";
      displayHikitori = displayHikitori || "不要";
    }
    return {
      kaisyacd: t.kaisyacd,
      kairname: t.kairname ?? "",
      coname: t.coname,
      hikitoriCurrent: displayHikitori,
      hikitoriDb: dbEx.hikitori,
      housyubeCurrent: displayHousyube,
      housyubeDb: dbEx.housyube,
      reportFlagMax: flag,
      needsReport,
    };
  });

  return { rows, truncated: tokus.length >= LIST_CAP };
}
