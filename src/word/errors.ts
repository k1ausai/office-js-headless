import { isSetSupported, SupportedPlatform } from "../office/context";

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

// Shared by Body.insertOoxml and Range.insertOoxml — real Word Online has no
// browser-side OOXML merge engine, a genuine, unfixed Microsoft product bug,
// not a missing capability (see design spec's "Platform selection" section).
export function assertOoxmlSupported(platform: SupportedPlatform): void {
  if (!isSetSupported(platform, "WordApiDesktop")) {
    throw richApiError(
      "GeneralException",
      "insertOoxml is not supported on this platform (Word Online has no client-side OOXML merge implementation)."
    );
  }
}
