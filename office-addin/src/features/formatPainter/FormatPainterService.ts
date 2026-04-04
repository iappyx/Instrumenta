/**
 * FormatPainterService.ts
 * Formatierungseigenschaften zwischen PowerPoint-Shapes kopieren.
 *
 * Workflow:
 * 1. captureFormat()     – liest Format des ersten selektierten Shapes
 * 2. applyFormat()       – wendet gespeichertes Format auf Ziel-Shapes an
 *
 * Unterstützte Eigenschaften:
 * - fill:     Solid Fill (Farbe + Transparenz). Gradient/Pattern: Hinweis, überspringen.
 * - line:     Farbe, Stärke, DashStyle, Transparenz.
 * - text:     Font-Name, -Größe, Fett, Kursiv, Farbe (TextRange-Level, nicht Run-Level).
 * - geometry: Adjustments (nur wenn Quell- und Ziel-Shape den gleichen geometrischen Typ haben).
 * - shadow:   Schatten-Eigenschaften (Office.js 1.5+, Mac-Support eingeschränkt, try/catch).
 *
 * Scope-Optionen:
 * - selection:      alle selektierten Shapes außer dem Quell-Shape
 * - slideByType:    alle Shapes des gleichen Shape-Typs auf der aktiven Slide
 * - deckByType:     alle Shapes des gleichen Shape-Typs auf allen Slides
 *
 * Office.js-Anforderungen:
 * - Shape.fill: 1.1+  |  Shape.lineFormat: 1.4+  |  Shape.textFrame: 1.1+
 * - Shape.geometricShape.adjustments: 1.4+
 * - Shape.shadow: 1.5+ (eingeschränkt auf Mac)
 */

import { logger }           from "../../utils/logger";
import { normalizeHex }     from "../../utils/colorUtils";
import { getSetting, setSetting, flushSettings } from "../../services/config/ConfigService";
import { pushUndo, createSnapshot } from "../../services/state/SessionState";

const MODULE   = "FormatPainterService";
const PRESETS_KEY = "INFRONT_FP_PRESETS";

// ─── Typen ─────────────────────────────────────────────────────────────────────

export type ApplyScope = "selection" | "slideByType" | "deckByType";

export interface FormatOptions {
  fill:     boolean;
  line:     boolean;
  text:     boolean;
  geometry: boolean;
  shadow:   boolean;
}

export interface FormatPreset {
  name:    string;
  options: FormatOptions;
}

export interface CapturedFill {
  color:        string | null;
  transparency: number;
  /** true wenn Gradient/Picture/Texture – dann nicht anwendbar */
  unsupported:  boolean;
}

export interface CapturedLine {
  color:        string | null;
  weight:       number;
  dashStyle:    string | null;
  transparency: number;
}

export interface CapturedText {
  fontName:  string | null;
  fontSize:  number | null;
  bold:      boolean;
  italic:    boolean;
  color:     string | null;
}

export interface CapturedFormat {
  shapeName:          string;
  shapeType:          string;
  geometricShapeType: string | null;
  fill:               CapturedFill   | null;
  line:               CapturedLine   | null;
  text:               CapturedText   | null;
  adjustments:        number[]       | null;
  hasShadow:          boolean;
}

export interface ApplyResult {
  applied:  number;
  skipped:  number;
  errors:   string[];
}

// ─── Capture ───────────────────────────────────────────────────────────────────

/**
 * Liest das Format des ersten selektierten Shapes.
 * Gibt null zurück wenn kein Shape selektiert.
 */
export async function captureFormat(): Promise<CapturedFormat | null> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length === 0) {
      logger.warn(MODULE, "captureFormat: Kein Shape selektiert.");
      return null;
    }

    const shape = selection.items[0];
    shape.load(["id", "name", "type", "geometricShapeType"]);
    await context.sync();

    const captured: CapturedFormat = {
      shapeName:          shape.name,
      shapeType:          String(shape.type),
      geometricShapeType: shape.geometricShapeType ? String(shape.geometricShapeType) : null,
      fill:               null,
      line:               null,
      text:               null,
      adjustments:        null,
      hasShadow:          false,
    };

    // ── Fill ──────────────────────────────────────────────────────────────────
    try {
      shape.fill.load(["type", "foregroundColor", "transparency"]);
      await context.sync();
      const fillType = shape.fill.type;
      if (fillType === PowerPoint.ShapeFillType.solid) {
        captured.fill = {
          color:        normalizeHex(shape.fill.foregroundColor ?? ""),
          transparency: shape.fill.transparency ?? 0,
          unsupported:  false,
        };
      } else if (fillType !== PowerPoint.ShapeFillType.noFill) {
        captured.fill = { color: null, transparency: 0, unsupported: true };
      }
    } catch (err) {
      logger.debug(MODULE, "Fill lesen fehlgeschlagen.", err);
    }

    // ── Line ──────────────────────────────────────────────────────────────────
    try {
      shape.lineFormat.load(["color", "weight", "dashStyle", "transparency"]);
      await context.sync();
      captured.line = {
        color:        normalizeHex(shape.lineFormat.color ?? ""),
        weight:       shape.lineFormat.weight ?? 0,
        dashStyle:    shape.lineFormat.dashStyle ? String(shape.lineFormat.dashStyle) : null,
        transparency: shape.lineFormat.transparency ?? 0,
      };
    } catch (err) {
      logger.debug(MODULE, "Line lesen fehlgeschlagen.", err);
    }

    // ── Text ──────────────────────────────────────────────────────────────────
    try {
      const font = shape.textFrame.textRange.font;
      font.load(["name", "size", "bold", "italic", "color"]);
      await context.sync();
      captured.text = {
        fontName:  font.name ?? null,
        fontSize:  font.size ?? null,
        bold:      font.bold ?? false,
        italic:    font.italic ?? false,
        color:     normalizeHex(font.color ?? ""),
      };
    } catch {
      /* kein TextFrame */
    }

    // ── Geometrie / Adjustments ───────────────────────────────────────────────
    try {
      const adjs = shape.geometricShape.adjustments;
      adjs.load("items");
      await context.sync();
      captured.adjustments = adjs.items.map((a) => {
        a.load("value");
        return a;
      });
      await context.sync();
      captured.adjustments = (captured.adjustments as unknown as PowerPoint.ShapeAdjustment[])
        .map((a) => a.value);
    } catch {
      /* keine Adjustments */
    }

    // ── Shadow ────────────────────────────────────────────────────────────────
    try {
      // shape.shadow ist in Office.js PowerPoint 1.5+ verfügbar,
      // Mac-Support ist eingeschränkt (try/catch als Absicherung).
      shape.load("shadow" as keyof PowerPoint.Shape);
      await context.sync();
      captured.hasShadow = true;
    } catch {
      captured.hasShadow = false;
    }

    logger.info(MODULE, `Format erfasst von: "${captured.shapeName}"`);
    return captured;
  }).catch((err) => {
    logger.error(MODULE, "captureFormat fehlgeschlagen.", err);
    return null;
  });
}

// ─── Apply ─────────────────────────────────────────────────────────────────────

/**
 * Wendet ein erfasstes Format auf Shapes an.
 *
 * @param format  - Erfasstes Format (von captureFormat())
 * @param options - Welche Eigenschaften übertragen werden
 * @param scope   - Auf welche Shapes anwenden
 */
export async function applyFormat(
  format:  CapturedFormat,
  options: FormatOptions,
  scope:   ApplyScope
): Promise<ApplyResult> {
  const result: ApplyResult = { applied: 0, skipped: 0, errors: [] };

  await PowerPoint.run(async (context) => {
    const targets = await resolveTargets(format, scope, context);

    if (targets.length === 0) {
      logger.warn(MODULE, "applyFormat: Keine Ziel-Shapes gefunden.");
      return;
    }

    // Undo-Snapshot
    const snapshots = targets.map((s) =>
      createSnapshot("", s.id, s.name, { options })
    );
    pushUndo({ featureName: "Format Painter+", timestamp: Date.now(), snapshots });

    for (const shape of targets) {
      try {
        await applySingleShape(shape, format, options, context);
        result.applied++;
      } catch (err) {
        logger.error(MODULE, `applyFormat fehlgeschlagen für "${shape.name}".`, err);
        result.errors.push(shape.name);
        result.skipped++;
      }
    }

    await context.sync();
  });

  return result;
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

/** Ermittelt alle Ziel-Shapes je nach Scope. */
async function resolveTargets(
  format:  CapturedFormat,
  scope:   ApplyScope,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape[]> {
  if (scope === "selection") {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items");
    await context.sync();
    // Quell-Shape (index 0) ausschließen
    const items = sel.items;
    for (const s of items) s.load(["id", "name", "type", "geometricShapeType"]);
    await context.sync();
    return items.slice(1);
  }

  if (scope === "slideByType") {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();
    if (slides.items.length === 0) return [];
    const slide = slides.items[0];
    return await getMatchingShapes(slide, format, context);
  }

  if (scope === "deckByType") {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    const all: PowerPoint.Shape[] = [];
    for (const slide of slides.items) {
      const matches = await getMatchingShapes(slide, format, context);
      all.push(...matches);
    }
    return all;
  }

  return [];
}

/** Gibt alle Shapes einer Slide zurück die zum gleichen Typ gehören. */
async function getMatchingShapes(
  slide:   PowerPoint.Slide,
  format:  CapturedFormat,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape[]> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  for (const s of shapes.items) s.load(["id", "name", "type", "geometricShapeType"]);
  await context.sync();
  return shapes.items.filter((s) =>
    String(s.type) === format.shapeType &&
    (format.geometricShapeType === null ||
     String(s.geometricShapeType) === format.geometricShapeType)
  );
}

/** Wendet das Format auf ein einzelnes Shape an. */
async function applySingleShape(
  shape:   PowerPoint.Shape,
  format:  CapturedFormat,
  options: FormatOptions,
  context: PowerPoint.RequestContext
): Promise<void> {
  // ── Fill ──────────────────────────────────────────────────────────────────
  if (options.fill && format.fill && !format.fill.unsupported && format.fill.color) {
    shape.fill.setSolidColor(format.fill.color);
    if (format.fill.transparency > 0) {
      shape.fill.transparency = format.fill.transparency;
    }
  }

  // ── Line ──────────────────────────────────────────────────────────────────
  if (options.line && format.line) {
    if (format.line.color)   shape.lineFormat.color        = format.line.color;
    if (format.line.weight)  shape.lineFormat.weight       = format.line.weight;
    if (format.line.transparency > 0) shape.lineFormat.transparency = format.line.transparency;
    if (format.line.dashStyle) {
      try {
        (shape.lineFormat as Record<string, unknown>)["dashStyle"] = format.line.dashStyle;
      } catch { /* dashStyle ggf. nicht setzbar */ }
    }
  }

  // ── Text ──────────────────────────────────────────────────────────────────
  if (options.text && format.text) {
    try {
      const font = shape.textFrame.textRange.font;
      if (format.text.fontName) font.name  = format.text.fontName;
      if (format.text.fontSize) font.size  = format.text.fontSize;
      font.bold   = format.text.bold;
      font.italic = format.text.italic;
      if (format.text.color) font.color = format.text.color;
    } catch { /* kein TextFrame */ }
  }

  // ── Geometrie / Adjustments ───────────────────────────────────────────────
  if (options.geometry && format.adjustments && format.adjustments.length > 0) {
    // Nur anwenden wenn gleicher geometrischer Typ
    shape.load("geometricShapeType");
    await context.sync();
    if (String(shape.geometricShapeType) === format.geometricShapeType) {
      try {
        const adjs = shape.geometricShape.adjustments;
        adjs.load("items");
        await context.sync();
        for (let i = 0; i < Math.min(adjs.items.length, format.adjustments.length); i++) {
          adjs.items[i].value = format.adjustments[i];
        }
      } catch { /* keine Adjustments */ }
    }
  }

  // ── Shadow ────────────────────────────────────────────────────────────────
  if (options.shadow && format.hasShadow) {
    // Shadow-API sehr eingeschränkt auf Mac; Implementierung in späteren Versionen
    logger.debug(MODULE, "Shadow-Transfer: noch nicht vollständig implementiert (Mac-Einschränkung).");
  }
}

// ─── Presets ──────────────────────────────────────────────────────────────────

/** Lädt gespeicherte Presets aus Document Settings. */
export function loadPresets(): FormatPreset[] {
  return getSetting<FormatPreset[]>(PRESETS_KEY, []);
}

/** Speichert einen neuen Preset. Überschreibt vorhandenen Preset mit gleichem Namen. */
export async function savePreset(preset: FormatPreset): Promise<void> {
  const presets = loadPresets().filter((p) => p.name !== preset.name);
  presets.push(preset);
  setSetting(PRESETS_KEY, presets);
  await flushSettings();
  logger.info(MODULE, `Preset gespeichert: "${preset.name}"`);
}

/** Löscht einen Preset nach Name. */
export async function deletePreset(name: string): Promise<void> {
  const presets = loadPresets().filter((p) => p.name !== name);
  setSetting(PRESETS_KEY, presets);
  await flushSettings();
}
