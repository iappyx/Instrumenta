/**
 * ColorPickerService.ts
 * Business-Logik für das Color-Picker-Feature.
 *
 * WICHTIG – Mac/Office.js-Einschränkung:
 * Ein systemweiter Screen-Pixel-Picker ist im Office Add-in auf Mac NICHT möglich.
 * WebKit/WKWebView (die Browser-Engine in PowerPoint für Mac) unterstützt die
 * EyeDropper API nicht (Stand 2026, nur Chrome 95+ / Edge).
 * → Fallback: Farbe aus selektiertem Shape auslesen + Hex/RGB-Eingabe + Markenfarben.
 *
 * Unterstützte Anwende-Ziele:
 * - fill  : shape.fill.setSolidColor() – nur bei Solid-Fill; Gradient/Pattern → Hinweis
 * - line  : shape.lineFormat.color – bei Shapes ohne Linie: kein Fehler, nur Zuweisung
 * - font  : shape.textFrame.textRange.font.color – nur bei Shapes mit TextFrame
 *
 * Office.js-Anforderungen:
 * - Shape.fill: PowerPoint API 1.1+
 * - Shape.lineFormat: PowerPoint API 1.4+
 * - Shape.textFrame: PowerPoint API 1.1+
 * - getSelectedShapes(): PowerPoint API 1.5+
 */

import { logger } from "../../utils/logger";
import { normalizeHex } from "../../utils/colorUtils";
import { pushUndo, createSnapshot } from "../../services/state/SessionState";

const MODULE = "ColorPickerService";

export type ColorTarget = "fill" | "line" | "font";

export interface ColorApplyResult {
  applied:  number;
  skipped:  number;
  errors:   string[];
}

export interface PickedColor {
  hex:    string;
  target: ColorTarget;
}

// ─── Session-Level Recent Colors ──────────────────────────────────────────────
// Wird im Modul-Scope gehalten (kein React-State), damit die Liste
// beim Wechsel zwischen Panels erhalten bleibt.

const MAX_RECENT = 8;
const recentColors: string[] = [];

/** Fügt eine Farbe zur zuletzt-verwendet-Liste hinzu. */
export function addRecentColor(hex: string): void {
  const norm = normalizeHex(hex);
  if (!norm) return;
  const idx = recentColors.indexOf(norm);
  if (idx !== -1) recentColors.splice(idx, 1);
  recentColors.unshift(norm);
  if (recentColors.length > MAX_RECENT) recentColors.length = MAX_RECENT;
}

/** Gibt die zuletzt verwendeten Farben zurück. */
export function getRecentColors(): string[] {
  return [...recentColors];
}

// ─── Farbe aus Shape lesen ────────────────────────────────────────────────────

/**
 * Liest die Farbe des ersten selektierten Shapes für das angegebene Ziel.
 * Gibt null zurück wenn kein Shape selektiert, das Ziel nicht lesbar oder
 * der Wert kein valider Hex-Code ist.
 */
export async function pickColorFromShape(target: ColorTarget): Promise<PickedColor | null> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length === 0) {
      logger.warn(MODULE, "pickColorFromShape: Kein Shape selektiert.");
      return null;
    }

    const shape = selection.items[0];

    try {
      if (target === "fill") {
        shape.fill.load("foregroundColor");
        await context.sync();
        const raw = shape.fill.foregroundColor;
        if (!raw) {
          logger.warn(MODULE, "Füllfarbe nicht lesbar (evtl. Gradient oder kein Fill).");
          return null;
        }
        const hex = normalizeHex(raw);
        return hex ? { hex, target } : null;
      }

      if (target === "line") {
        shape.lineFormat.load("color");
        await context.sync();
        const raw = shape.lineFormat.color;
        if (!raw) {
          logger.warn(MODULE, "Linienfarbe nicht lesbar.");
          return null;
        }
        const hex = normalizeHex(raw);
        return hex ? { hex, target } : null;
      }

      if (target === "font") {
        const tf = shape.textFrame;
        tf.textRange.font.load("color");
        await context.sync();
        const raw = tf.textRange.font.color;
        if (!raw) {
          logger.warn(MODULE, "Schriftfarbe nicht lesbar.");
          return null;
        }
        const hex = normalizeHex(raw);
        return hex ? { hex, target } : null;
      }
    } catch (err) {
      logger.error(MODULE, `pickColorFromShape(${target}) fehlgeschlagen.`, err);
    }

    return null;
  }).catch((err) => {
    logger.error(MODULE, "pickColorFromShape: PowerPoint.run fehlgeschlagen.", err);
    return null;
  });
}

// ─── Farbe auf Shapes anwenden ────────────────────────────────────────────────

/**
 * Wendet eine Farbe auf alle selektierten Shapes an.
 *
 * @param hex    - Hex-Farbe (#RRGGBB)
 * @param target - Ziel: Füllung, Linie oder Schrift
 */
export async function applyColorToShapes(
  hex:    string,
  target: ColorTarget
): Promise<ColorApplyResult> {
  const normHex = normalizeHex(hex);
  if (!normHex) {
    return { applied: 0, skipped: 0, errors: ["Ungültiger Hex-Wert."] };
  }

  const result: ColorApplyResult = { applied: 0, skipped: 0, errors: [] };

  await PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length === 0) {
      logger.warn(MODULE, "applyColorToShapes: Keine Shapes selektiert.");
      return;
    }

    const shapes = selection.items;

    // Batch-Load Name + ID für Snapshots
    for (const s of shapes) {
      s.load(["id", "name"]);
    }
    await context.sync();

    // Undo-Snapshot: alten Farbwert lesen
    const snapshots = await buildColorSnapshots(shapes, target, context);
    if (snapshots.length > 0) {
      pushUndo({ featureName: `Color Picker (${target})`, timestamp: Date.now(), snapshots });
    }

    // Farbe anwenden
    for (const shape of shapes) {
      try {
        if (target === "fill") {
          shape.fill.setSolidColor(normHex);
          result.applied++;
        } else if (target === "line") {
          shape.lineFormat.color = normHex;
          result.applied++;
        } else if (target === "font") {
          // Prüfen ob Shape einen TextFrame hat (defensive Zuweisung)
          shape.load("type");
          await context.sync();
          try {
            shape.textFrame.textRange.font.color = normHex;
            result.applied++;
          } catch {
            logger.debug(MODULE, `Shape "${shape.name}": kein TextFrame, übersprungen.`);
            result.skipped++;
          }
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        logger.error(MODULE, `Shape "${shape.name}" fehlgeschlagen: ${msg}`, err);
        result.errors.push(shape.name);
        result.skipped++;
      }
    }

    await context.sync();

    // Farbe zu "zuletzt verwendet" hinzufügen
    addRecentColor(normHex);

    logger.info(
      MODULE,
      `applyColorToShapes(${target}, ${normHex}): ${result.applied} angepasst, ` +
      `${result.skipped} übersprungen, ${result.errors.length} Fehler.`
    );
  });

  return result;
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

/** Liest die aktuellen Farbwerte aller Shapes für den Undo-Snapshot. */
async function buildColorSnapshots(
  shapes:  PowerPoint.Shape[],
  target:  ColorTarget,
  context: PowerPoint.RequestContext
) {
  const snapshots = [];
  for (const shape of shapes) {
    let oldColor: string | null = null;
    try {
      if (target === "fill") {
        shape.fill.load("foregroundColor");
        await context.sync();
        oldColor = shape.fill.foregroundColor;
      } else if (target === "line") {
        shape.lineFormat.load("color");
        await context.sync();
        oldColor = shape.lineFormat.color;
      } else if (target === "font") {
        shape.textFrame.textRange.font.load("color");
        await context.sync();
        oldColor = shape.textFrame.textRange.font.color;
      }
    } catch {
      // Lesefehler → Snapshot ohne Farbwert
    }
    snapshots.push(createSnapshot("", shape.id, shape.name, { color: oldColor, target }));
  }
  return snapshots;
}
