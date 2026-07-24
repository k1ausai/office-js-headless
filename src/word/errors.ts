// Shared shape for every error this shim throws that mirrors a real
// Office.js runtime error — traced from the shipped runtime, which aliases
// OfficeExtension.Error to an internal class whose .name is actually
// "RichApi.Error", not the documented-but-unobserved "OfficeExtension.Error"
// (see doc/wayfinder/tickets/001-error-message-shape.md and 007-error-name-value.md).
export function richApiError(code: string, message: string): Error {
  const error = new Error(message);
  error.name = "RichApi.Error";
  (error as Error & { code: string }).code = code;
  return error;
}
