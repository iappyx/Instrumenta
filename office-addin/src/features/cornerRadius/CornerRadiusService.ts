/**
 * CornerRadiusService.ts
 * Business-Logik für das Corner-Radius-Feature.
 *
 * Unterstützte Shape-Typen:
 * - roundedRectangle: adjustment[0] steuert den Eckenradius (normiert 0–1)
 *
 * Nicht unterstützt (Office.js API):
 * - snipRoundRect, snip1Rect etc.: andere Adjustment-Semantik, keine verlässliche
 *   Radius-Kontrolle über adjustment[0]
 * - Alle anderen Shape-Typen: kein Eckenradius-Adjustment vorhanden
 *
 * Pixel → Punkte: 1 px bei 96 DPI = 0,75 pt
 * Normierung: adjustmentValue = ptValue / (min(width, height) / 2), geklemmt [0, 1]
 *
 * Office.js-Anforderung: PowerPoint API 1.4+ (Shape.geometricShape.adjustments)
 */

import { logger } from "../../utils/logger";
import { pxToPt } from "../../utils/geometryUtils";
import { pushUndo, createSnapshot } from "../../services/state/SessionState";

const MODULE = "CornerRadiusService";

export interface CornerRadiusResult {
  applied:  number;
  skipped:  number;
  /** Shape-Namen, bei denen ein Fehler aufgetreten ist */
  errors:   string[];
}

/**
 * Wendet den Eckenradius auf alle selektierten roundedRectangle-Shapes an.
 *
 * @param radiusPx - Gewünschter Radius in Pixeln (≥ 0)
 * @returns Ergebnis mit Anzahl angepasster, übersprungener und fehlerhafter Shapes
 */
export async function applyCornerRadius(radiusPx: number): Promise<CornerRadiusResult> {
  const result: CornerRadiusResult = { applied: 0, skipped: 0, errors: [] };

  await PowerPoint.run(async (context) => {
    // Selektion laden
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length === 0) {
      logger.warn(MODULE, "Keine Shapes selektiert.");
      return;
    }

    // Schritt 1: Batch-Load geometricShapeType für alle selektierten Shapes
    const shapes = selection.items;
    for (const shape of shapes) {
      shape.load(["id", "name", "geometricShapeType", "width", "height"]);
    }
    await context.sync();

    // Schritt 2: Undo-Snapshot VOR der Änderung
    const snapshots = shapes
      .filter((s) => s.geometricShapeType === PowerPoint.GeometricShapeType.roundedRectangle)
      .map((s) => {
        // Bestehenden Adjustment-Wert lesen für Snapshot
        let oldValue = 0;
        try {
          const adj = s.geometricShape.adjustments.getItemAt(0);
          adj.load("value");
          oldValue = adj.value;
        } catch {
          // Adjustment nicht lesbar – Snapshot mit 0
        }
        return createSnapshot(
          /* slideId */  "",  // Slide-ID nicht kritisch für Undo-Anzeige
          /* shapeId */  s.id,
          /* shapeName */ s.name,
          { adjustmentValue: oldValue, radiusPx }
        );
      });

    if (snapshots.length > 0) {
      pushUndo({
        featureName: "Corner Radius",
        timestamp: Date.now(),
        snapshots,
      });
    }

    // Schritt 3: Radius anwenden (nur roundedRectangle)
    const ptValue = pxToPt(radiusPx);

    for (const shape of shapes) {
      const gst = shape.geometricShapeType;

      if (gst !== PowerPoint.GeometricShapeType.roundedRectangle) {
        logger.debug(MODULE, `Shape "${shape.name}" (Typ: ${gst}) übersprungen.`);
        result.skipped++;
        continue;
      }

      try {
        // Normierung: adjustment[0] ist im Bereich [0, 1]
        // 0 = kein Radius, 1 = maximaler Radius (min(w,h)/2)
        const maxRadiusPt = Math.min(shape.width, shape.height) / 2;
        const normalized  = maxRadiusPt > 0
          ? Math.min(ptValue / maxRadiusPt, 1.0)
          : 0;

        const adjustment = shape.geometricShape.adjustments.getItemAt(0);
        adjustment.value = normalized;

        result.applied++;
        logger.debug(
          MODULE,
          `Shape "${shape.name}": ${radiusPx}px → ${ptValue.toFixed(2)}pt → normalized=${normalized.toFixed(3)}`
        );
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        logger.error(MODULE, `Shape "${shape.name}" fehlgeschlagen: ${msg}`, err);
        result.errors.push(shape.name);
        result.skipped++;
      }
    }

    await context.sync();
    logger.info(
      MODULE,
      `Fertig: ${result.applied} angepasst, ${result.skipped} übersprungen, ${result.errors.length} Fehler.`
    );
  });

  return result;
}

/**
 * Berechnet den aktuellen Eckenradius (in px) des ersten selektierten roundedRectangle.
 * Nützlich für "aktuellen Wert anzeigen" in der UI.
 *
 * @returns Radius in Pixel oder null wenn nicht lesbar / kein roundedRectangle
 */
export async function readCurrentRadiusPx(): Promise<number | null> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length === 0) return null;

    const shape = selection.items[0];
    shape.load(["geometricShapeType", "width", "height"]);
    await context.sync();

    if (shape.geometricShapeType !== PowerPoint.GeometricShapeType.roundedRectangle) {
      return null;
    }

    try {
      const adj = shape.geometricShape.adjustments.getItemAt(0);
      adj.load("value");
      await context.sync();

      const maxRadiusPt = Math.min(shape.width, shape.height) / 2;
      const ptValue = adj.value * maxRadiusPt;
      // pt → px: 1 pt = 1/0.75 px
      return Math.round(ptValue / 0.75);
    } catch {
      return null;
    }
  }).catch((err) => {
    logger.warn(MODULE, "readCurrentRadiusPx fehlgeschlagen.", err);
    return null;
  });
}
