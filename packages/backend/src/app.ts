import cors from "cors";
import express, { type NextFunction, type Request, type Response } from "express";
import passport from "passport";

import { configurePassport } from "./auth/passport.js";
import { createSessionMiddleware } from "./config/session.js";
import { warnSampleDataModeOnce } from "./lib/sample-data-mode.js";
import { errorHandler } from "./middleware/error-handler.js";
import { notFoundHandler } from "./middleware/not-found.js";
import { apiRouter } from "./routes/index.js";
import { sseRouter } from "./sse/index.js";

export function createApp(): express.Express {
  warnSampleDataModeOnce();

  const app = express();
  const corsOrigin = process.env.CORS_ORIGIN ?? "http://localhost:3100";

  app.use(
    cors({
      origin: corsOrigin,
      credentials: true,
    }),
  );
  app.use(express.json());

  configurePassport(app);
  app.use(createSessionMiddleware());
  app.use(passport.initialize());
  app.use(passport.session());

  if (process.env.NODE_ENV !== "production") {
    app.use(devRequestLogger);
  }

  app.use("/api", apiRouter);
  app.use("/sse", sseRouter);
  app.use(notFoundHandler);
  app.use(errorHandler);

  return app;
}

function devRequestLogger(req: Request, res: Response, next: NextFunction): void {
  const startTime = Date.now();

  res.on("finish", () => {
    const elapsed = Date.now() - startTime;
    console.info(`${req.method} ${req.originalUrl} ${res.statusCode} ${elapsed}ms`);
  });

  next();
}
