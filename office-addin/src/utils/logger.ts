/**
 * logger.ts
 * Zentrales Logging für das Infront Toolkit.
 * Im Production-Build werden Debug-Logs unterdrückt.
 *
 * Remote Logging (optional):
 *   Nach Aufruf von enableRemoteLogging(baseUrl) werden alle log()-Aufrufe
 *   zusätzlich als fire-and-forget POST /log an die macOS Begleit-App gesendet.
 */

export type LogLevel = "debug" | "info" | "warn" | "error";

const PREFIX = "[InfrontToolkit]";

const isDev = typeof process !== "undefined" && process.env?.NODE_ENV === "development";

// Remote logging state (module-level, intentionally mutable)
let remoteBaseUrl: string | null = null;

function formatMessage(level: LogLevel, module: string, message: string): string {
  return `${PREFIX}[${module}] ${message}`;
}

/**
 * Aktiviert Remote Logging zur macOS Begleit-App.
 * Alle nachfolgenden info/warn/error Aufrufe werden zusätzlich per POST /log gesendet.
 * Fehler werden still ignoriert (fire-and-forget).
 *
 * @param baseUrl - z.B. "http://localhost:3000"
 */
export function enableRemoteLogging(baseUrl: string): void {
  remoteBaseUrl = baseUrl;
}

/** Deaktiviert Remote Logging (z.B. nach Verbindungsverlust). */
export function disableRemoteLogging(): void {
  remoteBaseUrl = null;
}

function sendRemoteLog(level: LogLevel, module: string, message: string): void {
  if (!remoteBaseUrl) return;
  const url = `${remoteBaseUrl}/log`;
  const body = JSON.stringify({
    level,
    module,
    message,
    timestamp: Date.now(),
  });
  // Fire-and-forget: Fehler werden still ignoriert
  fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body,
  }).catch(() => { /* intentionally silent */ });
}

export const logger = {
  debug(module: string, message: string, ...args: unknown[]): void {
    if (isDev) {
      console.debug(formatMessage("debug", module, message), ...args);
    }
    // Debug-Logs werden nicht remote gesendet
  },

  info(module: string, message: string, ...args: unknown[]): void {
    console.info(formatMessage("info", module, message), ...args);
    sendRemoteLog("info", module, message);
  },

  warn(module: string, message: string, ...args: unknown[]): void {
    console.warn(formatMessage("warn", module, message), ...args);
    sendRemoteLog("warn", module, message);
  },

  error(module: string, message: string, error?: unknown): void {
    const errorDetail = error instanceof Error
      ? `${error.message}\n${error.stack ?? ""}`
      : String(error ?? "");
    console.error(formatMessage("error", module, message), errorDetail);
    const remoteMessage = errorDetail ? `${message} — ${error instanceof Error ? error.message : errorDetail}` : message;
    sendRemoteLog("error", module, remoteMessage);
  },
};
