import { Router } from "express";
import passport from "passport";
import { z } from "zod";

import { requireAuth } from "../middleware/auth.js";

const loginBody = z.object({
  employeeId: z.string().min(1),
  password: z.string().min(1),
});

const authRouter = Router();

authRouter.post("/login", (req, res, next) => {
  const parsed = loginBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Invalid request body" });
    return;
  }

  passport.authenticate("local", (err: unknown, user: Express.User | false) => {
    if (err) {
      next(err);
      return;
    }
    if (!user) {
      res.status(401).json({ error: "社員番号またはパスワードが正しくありません" });
      return;
    }
    req.logIn(user, (loginErr) => {
      if (loginErr) {
        next(loginErr);
        return;
      }
      res.json({
        data: {
          employeeId: user.employeeId,
          role: user.role,
        },
      });
    });
  })(req, res, next);
});

authRouter.post("/logout", (req, res, next) => {
  req.logout((err) => {
    if (err) {
      next(err);
      return;
    }
    res.json({ data: { ok: true as const } });
  });
});

authRouter.get("/me", requireAuth, (req, res) => {
  const user = req.user!;
  res.json({
    data: {
      employeeId: user.employeeId,
      role: user.role,
    },
  });
});

export { authRouter };
