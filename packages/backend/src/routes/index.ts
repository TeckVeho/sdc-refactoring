import { Router } from "express";
import { authRouter } from "./auth.js";
import { healthRouter } from "./health.js";
import { irradiationRouter } from "./irradiation/index.js";
import { inventoryRouter } from "./inventory/index.js";
import { dosimetryRouter } from "./dosimetry/index.js";
import { reportsRouter } from "./reports/index.js";
import { adminRouter } from "./admin/index.js";
import { dashboardRouter } from "./dashboard.js";

const apiRouter = Router();

apiRouter.use("/health", healthRouter);
apiRouter.use("/auth", authRouter);
apiRouter.use("/dashboard", dashboardRouter);
apiRouter.use("/irradiation", irradiationRouter);
apiRouter.use("/inventory", inventoryRouter);
apiRouter.use("/dosimetry", dosimetryRouter);
apiRouter.use("/reports", reportsRouter);
apiRouter.use("/admin", adminRouter);

export { apiRouter };
