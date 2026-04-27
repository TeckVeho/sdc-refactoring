export type ApiSuccess<T> = { data: T; error?: undefined };
export type ApiError = { data?: undefined; error: string };
export type ApiResponse<T> = ApiSuccess<T> | ApiError;
