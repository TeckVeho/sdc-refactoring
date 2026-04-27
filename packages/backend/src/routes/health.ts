import { Router } from "express";

const healthRouter = Router();

healthRouter.get("/", (_req, res) => {
  res.json({ data: { ok: true as const } });
});

export { healthRouter };
