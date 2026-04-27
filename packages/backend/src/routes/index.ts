import { Router } from "express";

import { authRouter } from "./auth.js";
import { healthRouter } from "./health.js";

const apiRouter = Router();

apiRouter.use("/health", healthRouter);
apiRouter.use("/auth", authRouter);

export { apiRouter };
