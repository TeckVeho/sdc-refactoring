import { Router } from "express";

import { sampleDataApiMiddleware } from "../middleware/sample-api-stubs.js";
import { adminRouter } from "./admin/index.js";
import { authRouter } from "./auth.js";
import { dashboardRouter } from "./dashboard.js";
import { dosimetryRouter } from "./dosimetry/index.js";
import { healthRouter } from "./health.js";
import { inventoryRouter } from "./inventory/index.js";
import { irradiationRouter } from "./irradiation/index.js";
import { reportsRouter } from "./reports/index.js";

const apiRouter = Router();

apiRouter.use(sampleDataApiMiddleware);
apiRouter.use("/health", healthRouter);
apiRouter.use("/auth", authRouter);
apiRouter.use("/dashboard", dashboardRouter);
apiRouter.use("/irradiation", irradiationRouter);
apiRouter.use("/inventory", inventoryRouter);
apiRouter.use("/dosimetry", dosimetryRouter);
apiRouter.use("/reports", reportsRouter);
apiRouter.use("/admin", adminRouter);

export { apiRouter };
