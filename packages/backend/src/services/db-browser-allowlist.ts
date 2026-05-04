import type { PrismaClient } from "@prisma/client";

export type DbBrowserTableMeta = {
  id: string;
  label: string;
};

/** 参照のみ・書き込みなし。機密テーブルは含めない */
export const DB_BROWSER_TABLES: DbBrowserTableMeta[] = [
  { id: "tokumst", label: "取引先マスタ（tokumst）" },
  { id: "sehmst", label: "製品マスタ（sehmst）" },
  { id: "shainmst", label: "社員マスタ（shainmst）" },
  { id: "zaiko", label: "在庫（zaiko）" },
  { id: "zaikor", label: "在庫履歴（zaikor）" },
  { id: "syukar", label: "出荷実績（syukar）" },
  { id: "syouk1", label: "照射1号機（syouk1）" },
  { id: "syouk2", label: "照射2号機（syouk2）" },
  { id: "syoukj3", label: "照射3号機（syoukj3）" },
  { id: "senkind", label: "線種マスタ（senkind）" },
  { id: "keicode", label: "計算式コード（keicode）" },
  { id: "kcdcnvmst", label: "GM/EBコード変換（kcdcnvmst）" },
  { id: "tsyjmst", label: "注文マスタ（tsyjmst）" },
  { id: "ratetbl", label: "線量率テーブル（ratetbl）" },
  { id: "sengnr1", label: "センサ1（sengnr1）" },
  { id: "kyouj2", label: "挟込J2（kyouj2）" },
  { id: "sengnr3", label: "センサ3（sengnr3）" },
];

const MAX_ROWS = 500;

export async function queryBrowserTable(
  prisma: PrismaClient,
  tableId: string,
  opts: { take: number; skip: number; q?: string },
): Promise<unknown[]> {
  const take = Math.min(Math.max(opts.take, 1), MAX_ROWS);
  const skip = Math.max(opts.skip, 0);
  const q = opts.q?.trim();

  switch (tableId) {
    case "tokumst":
      return prisma.tokumst.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { kaisyacd: { contains: q } },
                { coname: { contains: q } },
                { kairname: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { kaisyacd: "asc" },
      });
    case "sehmst":
      return prisma.sehmst.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { kaisyacd: { contains: q } },
                { sehncd: { contains: q } },
                { seiname: { contains: q } },
              ],
            }
          : undefined,
        orderBy: [{ kaisyacd: "asc" }, { sehncd: "asc" }],
      });
    case "shainmst":
      return prisma.shainmst.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { shano: { contains: q } },
                { shaname: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { shano: "asc" },
      });
    case "zaiko":
      return prisma.zaiko.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { uno: { contains: q } },
                { kaisyacd: { contains: q } },
                { kainame: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { uno: "desc" },
      });
    case "zaikor":
      return prisma.zaikor.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { uno: { contains: q } },
                { kaisyacd: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { uno: "desc" },
      });
    case "syukar":
      return prisma.syukar.findMany({
        take,
        skip,
        where: q ? { kaisyacd: { contains: q } } : undefined,
        orderBy: { id: "desc" },
      });
    case "syouk1":
      return prisma.syouk1.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { uno: { contains: q } },
                { kainame: { contains: q } },
                { syono: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { uno: "desc" },
      });
    case "syouk2":
      return prisma.syouk2.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { uno: { contains: q } },
                { kainame: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { uno: "desc" },
      });
    case "syoukj3":
      return prisma.syoukj3.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { uno: { contains: q } },
                { kainame: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { uno: "desc" },
      });
    case "senkind":
      return prisma.senkind.findMany({
        take,
        skip,
        where: q ? { OR: [{ kindcd: { contains: q } }, { name: { contains: q } }] } : undefined,
        orderBy: { kindcd: "asc" },
      });
    case "keicode":
      return prisma.keicode.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { keicode: { contains: q } },
                { keisiki: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { keicode: "asc" },
      });
    case "kcdcnvmst":
      return prisma.kcdcnvmst.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { gammacd: { contains: q } },
                { ebcd: { contains: q } },
                { memo: { contains: q } },
              ],
            }
          : undefined,
        orderBy: [{ gammacd: "asc" }, { ebcd: "asc" }],
      });
    case "tsyjmst":
      return prisma.tsyjmst.findMany({
        take,
        skip,
        where: q
          ? {
              OR: [
                { chumoncd: { contains: q } },
                { kaisyacd: { contains: q } },
              ],
            }
          : undefined,
        orderBy: { chumoncd: "asc" },
      });
    case "ratetbl":
      return prisma.ratetbl.findMany({
        take,
        skip,
        where: q ? { ratekey: { contains: q } } : undefined,
        orderBy: { ratekey: "asc" },
      });
    case "sengnr1":
      return prisma.sengnr1.findMany({
        take,
        skip,
        where: q ? { OR: [{ sdate: { contains: q } }, { event: { contains: q } }] } : undefined,
        orderBy: { id: "desc" },
      });
    case "kyouj2":
      return prisma.kyouj2.findMany({
        take,
        skip,
        where: q ? { rectime: { contains: q } } : undefined,
        orderBy: { id: "desc" },
      });
    case "sengnr3":
      return prisma.sengnr3.findMany({
        take,
        skip,
        where: q ? { OR: [{ sdate: { contains: q } }, { event: { contains: q } }] } : undefined,
        orderBy: { id: "desc" },
      });
    default:
      throw new Error(`未対応テーブル: ${tableId}`);
  }
}
