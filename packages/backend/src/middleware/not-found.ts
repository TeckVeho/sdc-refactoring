import type { RequestHandler } from "express";
import type { ApiResponse } from "../types/api-response.js";

export const notFoundHandler: RequestHandler = (req, res) => {
  const body: ApiResponse<never> = {
    error: `Not Found: ${req.method} ${req.originalUrl}`,
  };

  res.status(404).json(body);
};
