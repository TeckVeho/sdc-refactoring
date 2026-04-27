import type { ErrorRequestHandler } from "express";
import type { ApiResponse } from "../types/api-response.js";

type HttpErrorLike = {
  status?: number;
  statusCode?: number;
  message?: string;
};

function resolveStatus(error: unknown): number {
  const maybeError = error as HttpErrorLike;
  const candidate = maybeError.status ?? maybeError.statusCode;

  if (typeof candidate === "number" && candidate >= 400 && candidate < 600) {
    return candidate;
  }

  return 500;
}

export const errorHandler: ErrorRequestHandler = (error, _req, res, _next) => {
  const status = resolveStatus(error);
  const message =
    status >= 500
      ? "Internal Server Error"
      : error instanceof Error
        ? error.message
        : "Unexpected Error";

  if (status >= 500) {
    console.error(error);
  }

  const body: ApiResponse<never> = { error: message };
  res.status(status).json(body);
};
