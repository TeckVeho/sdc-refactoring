import { Router } from "express";

import { requireAuth } from "../middleware/auth.js";
import { getMachine1MonitorData } from "../services/machine1-monitor.js";

const sseRouter = Router();

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

/** 1号機照会モニタ（getMachine1MonitorData 同一スナップショットを SSE で配信; 30秒間隔。フロントは主に 30s ポーリング） */
sseRouter.get("/machine1/monitor", requireAuth, async (req, res, next) => {
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  try {
    const first = await getMachine1MonitorData();
    res.write(`data: ${JSON.stringify({ data: first })}\n\n`);
  } catch (e) {
    next(e);
    return;
  }

  const interval = setInterval(() => {
    getMachine1MonitorData()
      .then((d) => {
        res.write(`data: ${JSON.stringify({ data: d })}\n\n`);
      })
      .catch((err) => {
        console.error(err);
        clearInterval(interval);
        res.end();
      });
  }, MACHINE1_MONITOR_MS);

  req.on("close", () => {
    clearInterval(interval);
  });
});

export { sseRouter };
