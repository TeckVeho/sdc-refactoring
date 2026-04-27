export type UserRole = "admin" | "manager" | "operator" | "viewer";

export type ApiSuccess<T> = { data: T; error?: undefined; meta?: { total?: number } };
export type ApiError = { data?: undefined; error: string; meta?: undefined };
export type ApiResponse<T> = ApiSuccess<T> | ApiError;

export function ok<T>(data: T, meta?: { total?: number }): ApiSuccess<T> {
  return meta ? { data, meta } : { data };
}

export function err(message: string): ApiError {
  return { error: message };
}
