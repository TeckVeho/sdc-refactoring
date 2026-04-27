import type { NextFunction, Request, Response } from "express";
import { Router } from "express";

import { requireAuth } from "../middleware/auth.js";
import { getMachine1MonitorData } from "../services/machine1-monitor.js";
import { getMachineStatusSnapshot } from "../services/machine-status-snapshot.js";

const sseRouter = Router();

/**
 * 機械モニタ系 SSE: 接続直後1回 + 一定間隔で `data: ${JSON.stringify({ data })}` 送信
 */
function streamJsonDataSse(
  res: Response,
  req: Request,
  next: NextFunction,
  load: () => Promise<unknown>,
  intervalMs: number,
): void {
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  load()
    .then((first) => {
      res.write(`data: ${JSON.stringify({ data: first })}\n\n`);
      const interval = setInterval(() => {
        load()
          .then((d) => {
            res.write(`data: ${JSON.stringify({ data: d })}\n\n`);
          })
          .catch((err) => {
            console.error(err);
            clearInterval(interval);
            res.end();
          });
      }, intervalMs);
      req.on("close", () => {
        clearInterval(interval);
      });
    })
    .catch((e) => {
      next(e);
    });
}

/** Phase 1: 接続維持とイベント形式の検証用（本番は装置モニタ用データに差し替え） */
sseRouter.get("/pulse", requireAuth, (req, res) => {
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  const send = (payload: Record<string, unknown>) => {
    res.write(`data: ${JSON.stringify(payload)}\n\n`);
  };

  send({ type: "connected", at: new Date().toISOString() });

  const intervalMs = Number.parseInt(process.env.SSE_PULSE_INTERVAL_MS ?? "30000", 10);
  const every = Number.isFinite(intervalMs) && intervalMs > 1000 ? intervalMs : 30000;

  const timer = setInterval(() => {
    send({ type: "pulse", at: new Date().toISOString() });
  }, every);

  req.on("close", () => {
    clearInterval(timer);
  });
});

const MACHINE1_MONITOR_MS = 30_000;
const MACHINE_STATUS_MS = 60_000;

/** 1号機照会モニタ（getMachine1MonitorData 同一スナップショットを SSE で配信; 30秒間隔。フロントは主に 30s ポーリング） */
sseRouter.get("/machine1/monitor", requireAuth, (req, res, next) => {
  streamJsonDataSse(res, req, next, () => getMachine1MonitorData(), MACHINE1_MONITOR_MS);
});

/** 装置運転状況（1〜3号機＋直近24hトレンド / Ex仕様#5: SSE 1分） */
sseRouter.get("/machine-status", requireAuth, (req, res, next) => {
  streamJsonDataSse(res, req, next, () => getMachineStatusSnapshot(), MACHINE_STATUS_MS);
});

export { sseRouter };
