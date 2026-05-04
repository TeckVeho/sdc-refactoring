import { Router } from "express";
import { z } from "zod";

import { prisma } from "../../lib/prisma.js";
import { requireAuth } from "../../middleware/auth.js";
import {
  DB_BROWSER_TABLES,
  queryBrowserTable,
} from "../../services/db-browser-allowlist.js";

const dbBrowserRouter = Router();

const querySchema = z.object({
  table: z.string().min(1),
  take: z.coerce.number().int().min(1).max(500).default(100),
  skip: z.coerce.number().int().min(0).default(0),
  q: z.string().optional(),
});

const postSchema = z.object({
  table: z.string().min(1),
  take: z.number().int().min(1).max(500).default(100),
  skip: z.number().int().min(0).default(0),
  q: z.string().optional(),
});

/** GET /api/admin/db-browser/meta */
dbBrowserRouter.get("/meta", requireAuth, (_req, res) => {
  res.json({ data: { tables: DB_BROWSER_TABLES, maxRows: 500 } });
});

/** GET /api/admin/db-browser?table=&take=&skip=&q= */
dbBrowserRouter.get("/", requireAuth, async (req, res) => {
  const parsed = querySchema.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid query" });
    return;
  }

  const id = parsed.data.table;
  if (!DB_BROWSER_TABLES.some((t) => t.id === id)) {
    res.status(400).json({ error: `テーブルが許可リストにありません: ${id}` });
    return;
  }

  try {
    const rows = await queryBrowserTable(prisma, id, {
      take: parsed.data.take,
      skip: parsed.data.skip,
      q: parsed.data.q,
    });

    const serialized = rows.map((r) =>
      JSON.parse(
        JSON.stringify(r, (_k, v) => (typeof v === "bigint" ? v.toString() : v)),
      ),
    );

    res.json({
      data: {
        table: id,
        take: parsed.data.take,
        skip: parsed.data.skip,
        rowCount: serialized.length,
        rows: serialized,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

/** POST /api/admin/db-browser（JSON ボディで検索 — Excel 用など） */
dbBrowserRouter.post("/query", requireAuth, async (req, res) => {
  const parsed = postSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid body" });
    return;
  }

  const id = parsed.data.table;
  if (!DB_BROWSER_TABLES.some((t) => t.id === id)) {
    res.status(400).json({ error: `テーブルが許可リストにありません: ${id}` });
    return;
  }

  try {
    const rows = await queryBrowserTable(prisma, id, {
      take: parsed.data.take,
      skip: parsed.data.skip,
      q: parsed.data.q,
    });

    const serialized = rows.map((r) =>
      JSON.parse(
        JSON.stringify(r, (_k, v) => (typeof v === "bigint" ? v.toString() : v)),
      ),
    );

    res.json({
      data: {
        table: id,
        take: parsed.data.take,
        skip: parsed.data.skip,
        rowCount: serialized.length,
        rows: serialized,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e instanceof Error ? e.message : "Server error" });
  }
});

export { dbBrowserRouter };
