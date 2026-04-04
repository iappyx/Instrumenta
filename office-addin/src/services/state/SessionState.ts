/**
 * SessionState.ts
 * Lokaler Session-State für Undo-Fallback und Feature-übergreifende Daten.
 *
 * Da Office.js kein natives programmatisches Undo für PowerPoint anbietet,
 * werden Snapshots vor destruktiven Operationen im Session-State gespeichert.
 * Dieser State lebt nur während der Add-in-Sitzung (kein persistenter Storage).
 *
 * HINWEIS: Der State wird zurückgesetzt wenn die Task Pane geschlossen wird.
 * Für persistente Daten → ConfigService / Document Settings.
 */

import { logger } from "../../utils/logger";

const MODULE = "SessionState";

/** Snapshot einer Shape-Eigenschaft vor einer Änderung. */
export interface ShapeSnapshot {
  slideId:  string;
  shapeId:  string;
  shapeName: string;
  /** Serialisierte Eigenschaften (je Feature unterschiedlich) */
  data:     Record<string, unknown>;
}

/** Ein Undo-Eintrag für eine Feature-Operation. */
export interface UndoEntry {
  featureName: string;
  timestamp:   number;
  snapshots:   ShapeSnapshot[];
}

// ─── Interner State ───────────────────────────────────────────────────────────

const MAX_UNDO_ENTRIES = 10;

let undoStack: UndoEntry[] = [];

/** Ausgewähltes Quell-Shape für Format Painter+ (Shape-ID). */
let formatPainterSourceId: string | null = null;

// ─── Undo-Stack ───────────────────────────────────────────────────────────────

/** Speichert einen Undo-Eintrag auf den Stack. */
export function pushUndo(entry: UndoEntry): void {
  undoStack.unshift(entry);
  if (undoStack.length > MAX_UNDO_ENTRIES) {
    undoStack = undoStack.slice(0, MAX_UNDO_ENTRIES);
  }
  logger.debug(MODULE, `Undo-Eintrag für "${entry.featureName}" gespeichert (Stack: ${undoStack.length}).`);
}

/** Gibt den neuesten Undo-Eintrag zurück (ohne ihn zu entfernen). */
export function peekUndo(): UndoEntry | null {
  return undoStack[0] ?? null;
}

/** Entfernt und gibt den neuesten Undo-Eintrag zurück. */
export function popUndo(): UndoEntry | null {
  const entry = undoStack.shift() ?? null;
  if (entry) {
    logger.debug(MODULE, `Undo-Eintrag für "${entry.featureName}" abgerufen.`);
  }
  return entry;
}

/** Gibt an ob Undo-Einträge vorhanden sind. */
export function canUndo(): boolean {
  return undoStack.length > 0;
}

/** Leert den Undo-Stack. */
export function clearUndo(): void {
  undoStack = [];
  logger.debug(MODULE, "Undo-Stack geleert.");
}

// ─── Format Painter State ─────────────────────────────────────────────────────

/** Speichert die ID des Quell-Shapes für Format Painter+. */
export function setFormatPainterSource(shapeId: string | null): void {
  formatPainterSourceId = shapeId;
  logger.debug(MODULE, `Format Painter Quelle: ${shapeId ?? "keine"}.`);
}

/** Gibt die ID des Quell-Shapes zurück. */
export function getFormatPainterSource(): string | null {
  return formatPainterSourceId;
}

// ─── Allgemeine Hilfsfunktion ──────────────────────────────────────────────────

/** Erstellt einen Snapshot-Eintrag für ein Shape. */
export function createSnapshot(
  slideId:   string,
  shapeId:   string,
  shapeName: string,
  data:      Record<string, unknown>
): ShapeSnapshot {
  return { slideId, shapeId, shapeName, data };
}
