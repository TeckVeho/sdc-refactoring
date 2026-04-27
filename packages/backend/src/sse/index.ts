import { Router } from "express";

import { requireAuth } from "../middleware/auth.js";

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

export { sseRouter };
