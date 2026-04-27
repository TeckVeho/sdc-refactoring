import { Router } from "express";
import { monitorRouter } from "./monitor.js";
import { planRouter } from "./plan.js";
import { machineStatusRouter } from "./machine-status.js";
import { productionRouter } from "./production.js";
import { workOrderRouter } from "./work-order.js";

const irradiationRouter = Router();

irradiationRouter.use("/machine1/monitor", monitorRouter);
irradiationRouter.use("/machine1/plan", planRouter);
irradiationRouter.use("/machine-status", machineStatusRouter);
irradiationRouter.use("/production", productionRouter);
irradiationRouter.use("/machine1/work-order", workOrderRouter);

export { irradiationRouter };
