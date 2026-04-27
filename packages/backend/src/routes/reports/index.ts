import { Router } from "express";
import { unissuedSearchRouter } from "./unissued-search.js";
import { ric2DoseShortageRouter } from "./ric2-dose-shortage.js";
import { ric3DoseShortageRouter } from "./ric3-dose-shortage.js";
import { ric3DoseShortageSumidenRouter } from "./ric3-dose-shortage-sumiden.js";
import { shipmentMethodRouter } from "./shipment-method.js";
import { gammaResultsRouter } from "./gamma-results.js";

const reportsRouter = Router();

reportsRouter.use("/unissued-search", unissuedSearchRouter);
reportsRouter.use("/ric2-dose-shortage", ric2DoseShortageRouter);
reportsRouter.use("/ric3-dose-shortage", ric3DoseShortageRouter);
reportsRouter.use("/ric3-dose-shortage-sumiden", ric3DoseShortageSumidenRouter);
reportsRouter.use("/shipment-method", shipmentMethodRouter);
reportsRouter.use("/gamma-results", gammaResultsRouter);

export { reportsRouter };
