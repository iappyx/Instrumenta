/**
 * GapEqualizerService.ts
 * Gap-Equalizer für das Infront Toolkit.
 *
 * Modi:
 * 1. "equal"   – gleichmäßig verteilen: äußere Shapes bleiben fixiert,
 *                innere werden verschoben sodass alle Abstände gleich sind.
 * 2. "fixed"   – fester Abstand: erstes Shape bleibt fixiert, alle weiteren
 *                werden mit dem angegebenen Abstand neu positioniert.
 * 3. "pack"    – dicht packen (Abstand = 0): wie "fixed" mit gap=0.
 *
 * Richtungen: horizontal (left/width), vertikal (top/height), beide.
 *
 * Sortierung: Shapes werden nach ihrer Position (left für H, top für V)
 * sortiert. Shapes mit identischer Position behalten relative Reihenfolge.
 *
 * Mindest-Selektion: 3 Shapes (bei "equal"), 2 Shapes (bei "fixed"/"pack").
 *
 * Undo: Session-State-Snapshot vor jeder Operation.
 *
 * Office.js-Anforderungen:
 * - getSelectedShapes(): 1.5+
 * - Shape.left / .top / .width / .height: 1.1+
 */

import { logger }       from "../../utils/logger";
import { pushUndo, createSnapshot } from "../../services/state/SessionState";
import { right, bottom, equalGap } from "../../utils/geometryUtils";

const MODULE = "GapEqualizerService";

// ─── Typen ────────────────────────────────────────────────────────────────────

export type GapDirection = "horizontal" | "vertical" | "both";
export type GapMode      = "equal" | "fixed" | "pack";

export interface GapOptions {
  direction: GapDirection;
  mode:      GapMode;
  /** Nur bei mode="fixed": gewünschter Abstand in pt. */
  fixedGap?: number;
}

export interface GapResult {
  /** Anzahl verschobener Shapes. */
  adjusted:     number;
  /** Berechneter Abstand (pt) – bei "equal" berechnet, sonst = fixedGap. */
  computedGapH: number;
  computedGapV: number;
  errors:       string[];
}

export interface GapPreview {
  shapeCount:   number;
  computedGapH: number;   // pt, NaN wenn nicht berechenbar
  computedGapV: number;
  valid:        boolean;
  message:      string;
}

interface ShapeData {
  shapeId:  string;
  slideId:  string;
  name:     string;
  left:     number;
  top:      number;
  width:    number;
  height:   number;
}

// ─── Vorschau (ohne zu schreiben) ─────────────────────────────────────────────

export async function previewGap(options: GapOptions): Promise<GapPreview> {
  let result: GapPreview = {
    shapeCount: 0,
    computedGapH: NaN,
    computedGapV: NaN,
    valid: false,
    message: "",
  };

  await PowerPoint.run(async (context) => {
    const data = await loadSelectedShapes(context);
    result.shapeCount = data.shapes.length;

    if (data.shapes.length < minShapes(options)) {
      result.message = `Bitte mindestens ${minShapes(options)} Shapes selektieren.`;
      return;
    }

    if (options.mode === "equal") {
      if (options.direction !== "vertical") {
        const sorted = [...data.shapes].sort((a, b) => a.left - b.left);
        const span   = right(sorted[sorted.length - 1]) - sorted[0].left;
        const total  = sorted.reduce((s, sh) => s + sh.width, 0);
        result.computedGapH = equalGap(span, total, sorted.length);
      }
      if (options.direction !== "horizontal") {
        const sorted = [...data.shapes].sort((a, b) => a.top - b.top);
        const span   = bottom(sorted[sorted.length - 1]) - sorted[0].top;
        const total  = sorted.reduce((s, sh) => s + sh.height, 0);
        result.computedGapV = equalGap(span, total, sorted.length);
      }
    } else {
      const g = options.mode === "pack" ? 0 : (options.fixedGap ?? 0);
      result.computedGapH = g;
      result.computedGapV = g;
    }

    result.valid   = true;
    result.message = buildPreviewMessage(options, result);
  });

  return result;
}

function minShapes(options: GapOptions): number {
  return options.mode === "equal" ? 3 : 2;
}

function buildPreviewMessage(options: GapOptions, preview: GapPreview): string {
  const parts: string[] = [];
  if (options.direction !== "vertical" && !isNaN(preview.computedGapH)) {
    parts.push(`H-Abstand: ${preview.computedGapH.toFixed(1)} pt`);
  }
  if (options.direction !== "horizontal" && !isNaN(preview.computedGapV)) {
    parts.push(`V-Abstand: ${preview.computedGapV.toFixed(1)} pt`);
  }
  return parts.length > 0 ? parts.join(" | ") : "—";
}

// ─── Gap angleichen ───────────────────────────────────────────────────────────

export async function equalizeGaps(options: GapOptions): Promise<GapResult> {
  const result: GapResult = { adjusted: 0, computedGapH: NaN, computedGapV: NaN, errors: [] };

  await PowerPoint.run(async (context) => {
    const data = await loadSelectedShapes(context);

    if (data.shapes.length < minShapes(options)) {
      result.errors.push(`Bitte mindestens ${minShapes(options)} Shapes selektieren.`);
      return;
    }

    // Undo-Snapshot
    const snapshots = data.shapes.map((s) =>
      createSnapshot(data.slideId, s.shapeId, s.name, {
        left: s.left, top: s.top, width: s.width, height: s.height,
      })
    );
    pushUndo({ featureName: "Gap Equalizer", timestamp: Date.now(), snapshots });

    // Horizontal
    if (options.direction !== "vertical") {
      const { gap, adjusted } = applyGapH(data.shapes, options, context);
      result.computedGapH = gap;
      result.adjusted    += adjusted;
    }

    // Vertikal
    if (options.direction !== "horizontal") {
      const { gap, adjusted } = applyGapV(data.shapes, options, context);
      result.computedGapV = gap;
      result.adjusted    += adjusted;
    }

    await context.sync();
    logger.info(MODULE, `equalizeGaps: ${result.adjusted} Shapes angepasst.`);
  });

  return result;
}

// ─── Horizontal ───────────────────────────────────────────────────────────────

function applyGapH(
  shapes:  ShapeData[],
  options: GapOptions,
  context: PowerPoint.RequestContext
): { gap: number; adjusted: number } {
  const sorted = [...shapes].sort((a, b) => a.left - b.left);
  let gap = 0;
  let adjusted = 0;

  if (options.mode === "equal") {
    const span  = right(sorted[sorted.length - 1]) - sorted[0].left;
    const total = sorted.reduce((s, sh) => s + sh.width, 0);
    gap = equalGap(span, total, sorted.length);

    // Nur innere Shapes verschieben (erstes + letztes bleiben)
    let cursor = sorted[0].left + sorted[0].width;
    for (let i = 1; i < sorted.length - 1; i++) {
      const newLeft = cursor + gap;
      if (Math.abs(sorted[i].left - newLeft) > 0.01) {
        setShapeLeft(sorted[i].shapeId, newLeft, context);
        adjusted++;
      }
      cursor = newLeft + sorted[i].width;
    }
  } else {
    gap = options.mode === "pack" ? 0 : (options.fixedGap ?? 0);
    // Erstes Shape bleibt, alle weiteren verschieben
    let cursor = sorted[0].left + sorted[0].width;
    for (let i = 1; i < sorted.length; i++) {
      const newLeft = cursor + gap;
      if (Math.abs(sorted[i].left - newLeft) > 0.01) {
        setShapeLeft(sorted[i].shapeId, newLeft, context);
        adjusted++;
      }
      cursor = newLeft + sorted[i].width;
    }
  }

  return { gap, adjusted };
}

// ─── Vertikal ─────────────────────────────────────────────────────────────────

function applyGapV(
  shapes:  ShapeData[],
  options: GapOptions,
  context: PowerPoint.RequestContext
): { gap: number; adjusted: number } {
  const sorted = [...shapes].sort((a, b) => a.top - b.top);
  let gap = 0;
  let adjusted = 0;

  if (options.mode === "equal") {
    const span  = bottom(sorted[sorted.length - 1]) - sorted[0].top;
    const total = sorted.reduce((s, sh) => s + sh.height, 0);
    gap = equalGap(span, total, sorted.length);

    let cursor = sorted[0].top + sorted[0].height;
    for (let i = 1; i < sorted.length - 1; i++) {
      const newTop = cursor + gap;
      if (Math.abs(sorted[i].top - newTop) > 0.01) {
        setShapeTop(sorted[i].shapeId, newTop, context);
        adjusted++;
      }
      cursor = newTop + sorted[i].height;
    }
  } else {
    gap = options.mode === "pack" ? 0 : (options.fixedGap ?? 0);
    let cursor = sorted[0].top + sorted[0].height;
    for (let i = 1; i < sorted.length; i++) {
      const newTop = cursor + gap;
      if (Math.abs(sorted[i].top - newTop) > 0.01) {
        setShapeTop(sorted[i].shapeId, newTop, context);
        adjusted++;
      }
      cursor = newTop + sorted[i].height;
    }
  }

  return { gap, adjusted };
}

// ─── Shapes laden ─────────────────────────────────────────────────────────────

async function loadSelectedShapes(
  context: PowerPoint.RequestContext
): Promise<{ shapes: ShapeData[]; slideId: string }> {
  const selection = context.presentation.getSelectedShapes();
  selection.load("items");
  await context.sync();

  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  slide.load("id");

  for (const shape of selection.items) {
    shape.load(["id", "name", "left", "top", "width", "height"]);
  }
  await context.sync();

  const shapes: ShapeData[] = selection.items.map((s) => ({
    shapeId: s.id,
    slideId: slide.id,
    name:    s.name,
    left:    s.left,
    top:     s.top,
    width:   s.width,
    height:  s.height,
  }));

  return { shapes, slideId: slide.id };
}

// ─── Shape-Schreib-Hilfsfunktionen ────────────────────────────────────────────

/**
 * Setzt die left-Position eines Shapes per ID.
 * Nutzt getItemById() damit kein zweites load/sync nötig ist.
 */
function setShapeLeft(shapeId: string, newLeft: number, context: PowerPoint.RequestContext): void {
  try {
    const slide  = context.presentation.getSelectedSlides().getItemAt(0);
    const shape  = slide.shapes.getItemById(shapeId);
    shape.left   = newLeft;
  } catch (err) {
    logger.warn(MODULE, `setShapeLeft ${shapeId}: ${err}`);
  }
}

function setShapeTop(shapeId: string, newTop: number, context: PowerPoint.RequestContext): void {
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shape = slide.shapes.getItemById(shapeId);
    shape.top   = newTop;
  } catch (err) {
    logger.warn(MODULE, `setShapeTop ${shapeId}: ${err}`);
  }
}
