import type { UserRole } from "@prisma/client";
import type { RequestHandler } from "express";

import { isRelaxedGuestAuth } from "../lib/dev-guest-auth.js";
import { isSampleDataMode } from "../lib/sample-data-mode.js";

/**
 * DB に存在しない仮ユーザー。サンプル API または開発用ゲストアクセス時に使用。
 */
const SAMPLE_MODE_USER = {
  id: "__sdc_sample_user__",
  employeeId: "sample",
  role: "admin" as UserRole,
} satisfies Express.User;

export const requireAuth: RequestHandler = (req, res, next) => {
  if (isSampleDataMode() || isRelaxedGuestAuth()) {
    if (!req.user) {
      req.user = SAMPLE_MODE_USER;
    }
    next();
    return;
  }
  if (!req.isAuthenticated?.() || !req.user) {
    res.status(401).json({ error: "Unauthorized" });
    return;
  }
  next();
};
