import { Router } from "express";
import { calendarRouter } from "./calendar.js";
import { shiftScheduleRouter } from "./shift-schedule.js";
import { priceSearchRouter } from "./price-search.js";
import { dbBrowserRouter } from "./db-browser.js";
import { companyCodeConvertRouter } from "./company-code-convert.js";
import { shipmentMethodVerifyRouter } from "./shipment-method-verify.js";
import { shipmentMethodSsRouter } from "./shipment-method-ss.js";

const adminRouter = Router();

adminRouter.use("/calendar", calendarRouter);
adminRouter.use("/shift-schedule", shiftScheduleRouter);
adminRouter.use("/price-search", priceSearchRouter);
adminRouter.use("/db-browser", dbBrowserRouter);
adminRouter.use("/company-code-convert", companyCodeConvertRouter);
adminRouter.use("/shipment-method-verify", shipmentMethodVerifyRouter);
adminRouter.use("/shipment-method-ss", shipmentMethodSsRouter);

export { adminRouter };
