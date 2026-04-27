import { Router } from "express";
import { ovenManagementRouter } from "./oven-management.js";
import { doseSearchRouter } from "./dose-search.js";
import { irradiationResultsRouter } from "./irradiation-results.js";
import { doseRateCalcRouter } from "./dose-rate-calc.js";
import { jmm60Router } from "./jmm60.js";
import { jmm90Router } from "./jmm90.js";

const dosimetryRouter = Router();

dosimetryRouter.use("/oven-management", ovenManagementRouter);
dosimetryRouter.use("/dose-search", doseSearchRouter);
dosimetryRouter.use("/irradiation-results", irradiationResultsRouter);
dosimetryRouter.use("/dose-rate-calc", doseRateCalcRouter);
dosimetryRouter.use("/jmm60", jmm60Router);
dosimetryRouter.use("/jmm90", jmm90Router);

export { dosimetryRouter };
