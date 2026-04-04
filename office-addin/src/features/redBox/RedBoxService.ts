/**
 * RedBoxService.ts
 * Red Box / Safe-Area-Begrenzung für das Infront Toolkit.
 *
 * Die „Red Box" ist ein transparentes Rahmen-Rechteck (INFRONT_REDBOX),
 * das die Safe Area einer Folie markiert. Sie dient als Druckhilfsrahmen
 * und Layoutreferenz.
 *
 * Verhalten:
 * - Einzige Red Box pro Folie (Name: INFRONT_REDBOX)
 * - Transparente Füllung, konfigurierbare Rahmenlinie
 * - Abstände (Margins) von allen vier Seiten konfigurierbar
 * - Slidegröße wird aus der Präsentation gelesen (API 1.1+)
 * - Konfiguration wird in Document.Settings gespeichert
 *
 * Office.js-Anforderungen:
 * - presentation.slideWidth / slideHeight:  1.1+
 * - shapes.addGeometricShape():             1.4+
 * - ShapeFill.clear():                      1.1+
 */

import { logger }         from "../../utils/logger";
import { normalizeHex }   from "../../utils/colorUtils";
import { getSetting, setSetting, flushSettings } from "../../services/config/ConfigService";

const MODULE       = "RedBoxService";
const SHAPE_NAME   = "INFRONT_REDBOX";
const SETTINGS_KEY = "redBoxConfig";

// ─── Typen ────────────────────────────────────────────────────────────────────

export interface RedBoxConfig {
  marginTop:    number;   // pt
  marginRight:  number;   // pt
  marginBottom: number;   // pt
  marginLeft:   number;   // pt
  color:        string;   // Hex-Farbe, z.B. "#FF0000"
  weight:       number;   // Linienstärke in pt
  lineDash:     "solid" | "dash" | "dot";
}

export interface RedBoxStatus {
  currentSlideHasBox: boolean;
  totalBoxes:         number;
  slideWidth:         number;
  slideHeight:        number;
}

export interface RedBoxResult {
  added:   number;
  removed: number;
  errors:  string[];
}

// ─── Standard-Konfiguration ───────────────────────────────────────────────────

export const DEFAULT_REDBOX_CONFIG: RedBoxConfig = {
  marginTop:    20,
  marginRight:  20,
  marginBottom: 20,
  marginLeft:   20,
  color:        "#FF0000",
  weight:       1.5,
  lineDash:     "solid",
};

// ─── Konfiguration laden / speichern ──────────────────────────────────────────

export function loadRedBoxConfig(): RedBoxConfig {
  const saved = getSetting<Partial<RedBoxConfig> | null>(SETTINGS_KEY, null);
  return saved ? { ...DEFAULT_REDBOX_CONFIG, ...saved } : { ...DEFAULT_REDBOX_CONFIG };
}

export async function saveRedBoxConfig(config: RedBoxConfig): Promise<void> {
  setSetting(SETTINGS_KEY, config);
  await flushSettings();
  logger.info(MODULE, "saveRedBoxConfig: gespeichert.");
}

// ─── Slide-Größe lesen ────────────────────────────────────────────────────────

export async function getSlideSize(): Promise<{ width: number; height: number }> {
  let width  = 720;
  let height = 540;

  await PowerPoint.run(async (context) => {
    context.presentation.load(["slideWidth", "slideHeight"]);
    await context.sync();
    width  = context.presentation.slideWidth  || 720;
    height = context.presentation.slideHeight || 540;
  });

  return { width, height };
}

// ─── Status der aktuellen Folie ───────────────────────────────────────────────

export async function getRedBoxStatus(): Promise<RedBoxStatus> {
  const status: RedBoxStatus = {
    currentSlideHasBox: false,
    totalBoxes: 0,
    slideWidth:  720,
    slideHeight: 540,
  };

  await PowerPoint.run(async (context) => {
    context.presentation.load(["slideWidth", "slideHeight"]);
    const allSlides = context.presentation.slides;
    allSlides.load("items");
    await context.sync();

    status.slideWidth  = context.presentation.slideWidth  || 720;
    status.slideHeight = context.presentation.slideHeight || 540;

    // Aktuelle Folie
    const selected = context.presentation.getSelectedSlides();
    selected.load("items");
    await context.sync();

    if (selected.items.length > 0) {
      const currentSlide = selected.items[0];
      currentSlide.shapes.load("items/name");
      await context.sync();
      status.currentSlideHasBox = currentSlide.shapes.items.some(
        (s) => s.name === SHAPE_NAME
      );
    }

    // Alle Folien zählen
    for (const slide of allSlides.items) {
      slide.shapes.load("items/name");
    }
    await context.sync();

    for (const slide of allSlides.items) {
      if (slide.shapes.items.some((s) => s.name === SHAPE_NAME)) {
        status.totalBoxes++;
      }
    }
  });

  return status;
}

// ─── Red Box auf einer Folie einfügen ─────────────────────────────────────────

async function insertRedBoxOnSlide(
  slide:   PowerPoint.Slide,
  config:  RedBoxConfig,
  slideW:  number,
  slideH:  number,
  context: PowerPoint.RequestContext
): Promise<void> {
  const color = normalizeHex(config.color) ?? "#FF0000";

  const left   = config.marginLeft;
  const top    = config.marginTop;
  const width  = slideW - config.marginLeft - config.marginRight;
  const height = slideH - config.marginTop  - config.marginBottom;

  if (width <= 0 || height <= 0) {
    throw new Error("Ungültige Margins: Breite oder Höhe ≤ 0.");
  }

  const box = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left, top, width, height,
  });

  box.name = SHAPE_NAME;
  box.fill.clear();   // Transparent – kein Hintergrund
  box.lineFormat.color  = color;
  box.lineFormat.weight = config.weight;

  // Linientyp (solid ist default; dash/dot erfordern DashStyle-Enum)
  if (config.lineDash !== "solid") {
    try {
      box.lineFormat.dashStyle = config.lineDash === "dash"
        ? PowerPoint.ShapeLineDashStyle.dash
        : PowerPoint.ShapeLineDashStyle.roundDot;
    } catch { /* DashStyle API nicht verfügbar – bleibt solid */ }
  }

  try { box.tags.add("INFRONT_REDBOX_VERSION", "1"); } catch { /* */ }

  void context;
}

// ─── Toggle auf der aktuellen Folie ──────────────────────────────────────────

export async function toggleRedBoxOnCurrentSlide(config?: RedBoxConfig): Promise<"added" | "removed"> {
  const cfg = config ?? loadRedBoxConfig();
  let action: "added" | "removed" = "added";

  await PowerPoint.run(async (context) => {
    context.presentation.load(["slideWidth", "slideHeight"]);
    const selected = context.presentation.getSelectedSlides();
    selected.load("items");
    await context.sync();

    const slideW = context.presentation.slideWidth  || 720;
    const slideH = context.presentation.slideHeight || 540;

    if (selected.items.length === 0) throw new Error("Keine Folie selektiert.");
    const slide = selected.items[0];
    slide.shapes.load("items/name");
    await context.sync();

    const existing = slide.shapes.items.find((s) => s.name === SHAPE_NAME);
    if (existing) {
      existing.delete();
      action = "removed";
    } else {
      await insertRedBoxOnSlide(slide, cfg, slideW, slideH, context);
      action = "added";
    }

    await context.sync();
  });

  logger.info(MODULE, `toggleRedBoxOnCurrentSlide: ${action}.`);
  return action;
}

// ─── Red Box auf allen Folien einfügen ────────────────────────────────────────

export async function addRedBoxToAllSlides(config?: RedBoxConfig): Promise<RedBoxResult> {
  const cfg    = config ?? loadRedBoxConfig();
  const result = { added: 0, removed: 0, errors: [] as string[] };

  await PowerPoint.run(async (context) => {
    context.presentation.load(["slideWidth", "slideHeight"]);
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const slideW = context.presentation.slideWidth  || 720;
    const slideH = context.presentation.slideHeight || 540;

    for (const slide of slides.items) {
      slide.shapes.load("items/name");
    }
    await context.sync();

    for (const slide of slides.items) {
      const existing = slide.shapes.items.find((s) => s.name === SHAPE_NAME);
      if (existing) continue;   // bereits vorhanden → überspringen
      try {
        await insertRedBoxOnSlide(slide, cfg, slideW, slideH, context);
        result.added++;
      } catch (err) {
        result.errors.push(err instanceof Error ? err.message : String(err));
      }
    }

    await context.sync();
  });

  logger.info(MODULE, `addRedBoxToAllSlides: ${result.added} eingefügt, ${result.errors.length} Fehler.`);
  return result;
}

// ─── Red Box von allen Folien entfernen ───────────────────────────────────────

export async function removeRedBoxFromAllSlides(): Promise<RedBoxResult> {
  const result = { added: 0, removed: 0, errors: [] as string[] };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.shapes.load("items/name");
    }
    await context.sync();

    for (const slide of slides.items) {
      for (const shape of slide.shapes.items) {
        if (shape.name === SHAPE_NAME) {
          shape.delete();
          result.removed++;
        }
      }
    }

    await context.sync();
  });

  logger.info(MODULE, `removeRedBoxFromAllSlides: ${result.removed} entfernt.`);
  return result;
}

// ─── Red Box auf aktueller Folie aktualisieren ────────────────────────────────

export async function updateRedBoxOnCurrentSlide(config: RedBoxConfig): Promise<boolean> {
  let updated = false;

  await PowerPoint.run(async (context) => {
    context.presentation.load(["slideWidth", "slideHeight"]);
    const selected = context.presentation.getSelectedSlides();
    selected.load("items");
    await context.sync();

    const slideW = context.presentation.slideWidth  || 720;
    const slideH = context.presentation.slideHeight || 540;

    if (selected.items.length === 0) return;
    const slide = selected.items[0];
    slide.shapes.load("items/name");
    await context.sync();

    // Alte Box löschen + neue einfügen
    const existing = slide.shapes.items.find((s) => s.name === SHAPE_NAME);
    if (existing) {
      existing.delete();
      await context.sync();
    }

    await insertRedBoxOnSlide(slide, config, slideW, slideH, context);
    await context.sync();
    updated = true;
  });

  return updated;
}
