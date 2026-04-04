/**
 * ConfigService.ts
 * Lese- und Schreiboperationen für die Infront Toolkit Konfiguration.
 *
 * Strategie:
 * - Konfiguration wird in Office.context.document.settings gespeichert.
 * - Diese Settings sind pro Dokument persistent (auch nach Schließen/Öffnen).
 * - Für deck-unabhängige Einstellungen (z.B. Brand-Profil-Auswahl) wird
 *   localStorage als Fallback verwendet.
 *
 * HINWEIS: Office.context.document.settings.saveAsync ist asynchron und
 * wird nicht garantiert sofort auf Mac persistiert. Kritische Settings
 * immer mit saveAsync() flushen.
 */

import { BrandConfig, DEFAULT_BRAND_CONFIG } from "./BrandConfig";
import { logger } from "../../utils/logger";

const MODULE        = "ConfigService";
const BRAND_KEY     = "INFRONT_BRAND_CONFIG";
const REDBOX_KEY    = "INFRONT_REDBOX_CONFIG";
const AGENDA_KEY    = "INFRONT_AGENDA_CONFIG";

// ─── Brand Config ─────────────────────────────────────────────────────────────

/** Lädt die Brand-Konfiguration aus Document Settings (Fallback: Default). */
export function loadBrandConfig(): BrandConfig {
  try {
    const raw = Office.context.document.settings.get(BRAND_KEY) as string | null;
    if (raw) {
      return JSON.parse(raw) as BrandConfig;
    }
  } catch (err) {
    logger.warn(MODULE, "Brand-Config konnte nicht geladen werden, nutze Default.", err);
  }
  return DEFAULT_BRAND_CONFIG;
}

/** Speichert die Brand-Konfiguration in Document Settings. */
export async function saveBrandConfig(config: BrandConfig): Promise<void> {
  try {
    Office.context.document.settings.set(BRAND_KEY, JSON.stringify(config));
    await new Promise<void>((resolve, reject) => {
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? "Settings save failed"));
        }
      });
    });
    logger.info(MODULE, "Brand-Config gespeichert.");
  } catch (err) {
    logger.error(MODULE, "Brand-Config konnte nicht gespeichert werden.", err);
    throw err;
  }
}

// ─── Generischer Key-Value Store ──────────────────────────────────────────────

/** Liest einen beliebigen Wert aus den Document Settings. */
export function getSetting<T>(key: string, defaultValue: T): T {
  try {
    const raw = Office.context.document.settings.get(key) as T | null;
    return raw !== null && raw !== undefined ? raw : defaultValue;
  } catch {
    return defaultValue;
  }
}

/** Setzt einen beliebigen Wert in den Document Settings (ohne flush). */
export function setSetting(key: string, value: unknown): void {
  try {
    Office.context.document.settings.set(key, value);
  } catch (err) {
    logger.error(MODULE, `setSetting("${key}") fehlgeschlagen.`, err);
  }
}

/** Flusht alle ausstehenden Document Settings auf Disk. */
export async function flushSettings(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message ?? "Settings flush failed"));
      }
    });
  });
}

// ─── macOS Companion App Integration ─────────────────────────────────────────

const MAC_APP_PORT_KEY = "INFRONT_MAC_APP_PORT";
const DEFAULT_MAC_APP_PORT = 3000;

/** Gibt die Base-URL der lokalen macOS Begleit-App zurück (z.B. "http://localhost:3000"). */
export function getMacAppBaseUrl(): string {
  const port = getSetting<number>(MAC_APP_PORT_KEY, DEFAULT_MAC_APP_PORT);
  return `http://localhost:${port}`;
}

/**
 * Versucht die Brand-Konfiguration vom lokalen macOS Companion Server zu laden.
 * Fällt bei Fehler (App nicht gestartet, Netzwerkfehler) still auf null zurück.
 * Der Aufrufer entscheidet, ob der Fallback auf DEFAULT_BRAND_CONFIG erfolgt.
 */
export async function fetchBrandConfigFromServer(): Promise<BrandConfig | null> {
  const baseUrl = getMacAppBaseUrl();
  try {
    const response = await fetch(`${baseUrl}/config/Infront_BrandConfig.json`, {
      method: "GET",
      headers: { Accept: "application/json" },
      // Kurzes Timeout: macOS App ist lokal, sollte schnell antworten
      signal: AbortSignal.timeout(3000),
    });
    if (!response.ok) {
      logger.warn(MODULE, `macOS App Config-Fetch fehlgeschlagen: HTTP ${response.status}`);
      return null;
    }
    const config = (await response.json()) as BrandConfig;
    logger.info(MODULE, "Brand-Config vom macOS Companion geladen.");
    return config;
  } catch {
    // App nicht gestartet oder nicht erreichbar – stilles Fallback
    return null;
  }
}

// ─── Öffentliche Keys ─────────────────────────────────────────────────────────
export { BRAND_KEY, REDBOX_KEY, AGENDA_KEY };
