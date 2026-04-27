import type { Express } from "express";
import passport from "passport";
import { Strategy as LocalStrategy } from "passport-local";
import bcrypt from "bcryptjs";

import { prisma } from "../lib/prisma.js";

export function configurePassport(_app: Express): void {
  passport.use(
    new LocalStrategy(
      { usernameField: "employeeId", passwordField: "password" },
      async (employeeId, password, done) => {
        try {
          const shain = await prisma.shainmst.findUnique({ where: { shano: employeeId } });
          if (!shain) {
            done(null, false);
            return;
          }

          const user = await prisma.user.findUnique({ where: { employeeId } });
          if (!user) {
            done(null, false);
            return;
          }

          const ok = await bcrypt.compare(password, user.passwordHash);
          if (!ok) {
            done(null, false);
            return;
          }

          done(null, { id: user.id, employeeId: user.employeeId, role: user.role });
        } catch (err) {
          done(err as Error);
        }
      },
    ),
  );

  passport.serializeUser((user, done) => {
    done(null, user.id);
  });

  passport.deserializeUser(async (id: string, done) => {
    try {
      const user = await prisma.user.findUnique({ where: { id } });
      if (!user) {
        done(null, false);
        return;
      }
      done(null, { id: user.id, employeeId: user.employeeId, role: user.role });
    } catch (err) {
      done(err);
    }
  });
}
