/**
 * notifications.ts
 * Hilfsfunktionen für Benutzer-Benachrichtigungen.
 *
 * HINWEIS: Office.js bietet im Task-Pane-Kontext keine eigene Notification-API.
 * Benachrichtigungen werden über den NotificationBar-React-State in den
 * jeweiligen Panel-Komponenten realisiert.
 *
 * Diese Datei stellt typisierte Hilfstypen und Factory-Funktionen bereit.
 */

export type NotificationType = "success" | "warning" | "error" | "info";

export interface AppNotification {
  message: string;
  type:    NotificationType;
}

/** Erstellt eine Erfolgs-Benachrichtigung. */
export function success(message: string): AppNotification {
  return { message, type: "success" };
}

/** Erstellt eine Warn-Benachrichtigung. */
export function warning(message: string): AppNotification {
  return { message, type: "warning" };
}

/** Erstellt eine Fehler-Benachrichtigung. */
export function error(message: string): AppNotification {
  return { message, type: "error" };
}

/** Erstellt eine Info-Benachrichtigung. */
export function info(message: string): AppNotification {
  return { message, type: "info" };
}

/**
 * Formatiert eine Ergebnis-Meldung mit Anzahl betroffener / übersprungener Shapes.
 */
export function resultMessage(applied: number, skipped: number, unit = "Shape(s)"): AppNotification {
  if (applied === 0 && skipped === 0) {
    return warning("Keine Shapes gefunden.");
  }
  if (applied === 0) {
    return warning(`0 ${unit} angepasst (${skipped} übersprungen – Shape-Typ nicht unterstützt).`);
  }
  const skipNote = skipped > 0 ? `, ${skipped} übersprungen` : "";
  return success(`${applied} ${unit} angepasst${skipNote}.`);
}
