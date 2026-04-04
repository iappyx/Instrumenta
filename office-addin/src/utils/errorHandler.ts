/**
 * errorHandler.ts
 * Zentrales Error-Handling für Office.js-Fehler und allgemeine Laufzeitfehler.
 */

import { logger } from "./logger";

/** Bekannte Office.js-Fehlercodes mit deutschen Beschreibungen. */
const OFFICE_ERROR_MESSAGES: Record<string, string> = {
  "GeneralException":           "Allgemeiner PowerPoint-Fehler.",
  "ItemNotFound":               "Das angeforderte Element wurde nicht gefunden.",
  "AccessDenied":               "Zugriff verweigert.",
  "InvalidArgument":            "Ungültiger Wert übergeben.",
  "InvalidOperation":           "Operation nicht zulässig.",
  "ItemAlreadyExists":          "Element existiert bereits.",
  "UnsupportedApiByRuntime":    "Diese API wird von der aktuellen Office-Version nicht unterstützt.",
  "ApiNotFound":                "API nicht gefunden – bitte Office-Version prüfen.",
  "NotImplemented":             "Diese Funktion ist noch nicht implementiert.",
  "RequestPayloadSizeLimitExceeded": "Anfrage zu groß – bitte weniger Shapes selektieren.",
};

/**
 * Wandelt einen unbekannten Error-Wert in eine benutzerfreundliche deutsche Meldung um.
 */
export function toUserMessage(err: unknown): string {
  if (err instanceof Error) {
    // Office.js-Fehler haben oft einen `code`-Property
    const code = (err as Error & { code?: string }).code;
    if (code && OFFICE_ERROR_MESSAGES[code]) {
      return OFFICE_ERROR_MESSAGES[code];
    }
    return err.message || "Unbekannter Fehler.";
  }
  return String(err ?? "Unbekannter Fehler.");
}

/**
 * Loggt einen Fehler und gibt eine benutzerfreundliche Meldung zurück.
 */
export function handleError(module: string, context: string, err: unknown): string {
  const userMessage = toUserMessage(err);
  logger.error(module, `${context}: ${userMessage}`, err);
  return userMessage;
}

/**
 * Prüft, ob eine Office.js-API-Funktion in der aktuellen Umgebung verfügbar ist.
 * Nützlich für API-Version-Checks auf Mac.
 */
export function isApiAvailable(requirement: string, version: string): boolean {
  return Office.context.requirements.isSetSupported(requirement, version);
}
