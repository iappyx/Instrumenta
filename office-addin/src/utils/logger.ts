/**
 * logger.ts
 * Zentrales Logging für das Infront Toolkit.
 * Im Production-Build werden Debug-Logs unterdrückt.
 */

export type LogLevel = "debug" | "info" | "warn" | "error";

const PREFIX = "[InfrontToolkit]";

const isDev = typeof process !== "undefined" && process.env?.NODE_ENV === "development";

function formatMessage(level: LogLevel, module: string, message: string): string {
  return `${PREFIX}[${module}] ${message}`;
}

export const logger = {
  debug(module: string, message: string, ...args: unknown[]): void {
    if (isDev) {
      console.debug(formatMessage("debug", module, message), ...args);
    }
  },

  info(module: string, message: string, ...args: unknown[]): void {
    console.info(formatMessage("info", module, message), ...args);
  },

  warn(module: string, message: string, ...args: unknown[]): void {
    console.warn(formatMessage("warn", module, message), ...args);
  },

  error(module: string, message: string, error?: unknown): void {
    const errorDetail = error instanceof Error
      ? `${error.message}\n${error.stack ?? ""}`
      : String(error ?? "");
    console.error(formatMessage("error", module, message), errorDetail);
  },
};
