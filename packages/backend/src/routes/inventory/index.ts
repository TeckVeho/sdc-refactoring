import { Router } from "express";
import { shipmentSummaryRouter } from "./shipment-summary.js";
import { companyArrivalRouter } from "./company-arrival.js";
import { monthEndRouter } from "./month-end.js";
import { customerStockRouter } from "./customer-stock.js";
import { ric3RepackRouter } from "./ric3-repack.js";

const inventoryRouter = Router();

inventoryRouter.use("/shipment-summary", shipmentSummaryRouter);
inventoryRouter.use("/company-arrival", companyArrivalRouter);
inventoryRouter.use("/month-end", monthEndRouter);
inventoryRouter.use("/customer-stock", customerStockRouter);
inventoryRouter.use("/ric3-repack", ric3RepackRouter);

export { inventoryRouter };
