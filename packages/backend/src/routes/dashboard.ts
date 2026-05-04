import { Router } from "express";

import { requireAuth } from "../middleware/auth.js";

const dashboardRouter = Router();

/** 将来 DB 化するまでは固定。未設定時は空（フロントは非表示） */
const NOTICE_TEXT = "";

dashboardRouter.get("/notice", requireAuth, (_req, res) => {
  res.json({ data: { text: NOTICE_TEXT } });
});

export { dashboardRouter };
