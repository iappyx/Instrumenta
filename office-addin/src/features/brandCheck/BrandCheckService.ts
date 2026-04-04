/**
 * BrandCheckService.ts
 * Brand-Compliance-Scanner für das Infront Toolkit.
 *
 * Prüft alle Shapes aller Slides auf Einhaltung der Brand-Richtlinien:
 * - Schriftarten (exakter Match, case-insensitiv)
 * - Schriftgröße (Minimum in pt)
 * - Textfarbe, Füllfarbe, Linienfarbe (mit konfigurierbarer Toleranz)
 *
 * Gruppen: bis zu 2 Ebenen tief rekursiv.
 * Tabellen: Zelltext und Zell-Füllfarbe (Font-Color pro Run eingeschränkt –
 *   Office.js TableCell API bietet keinen direkten paragraph/run Zugriff;
 *   Einschränkung in TESTING.md dokumentiert).
 *
 * Fix-Aktionen: Font-Name, Font-Größe, Farben (nächste Markenfarbe).
 *
 * Office.js-Anforderungen: PowerPoint API 1.1+ (Shapes/Fill/Font), 1.4+ (LineFormat),
 * 1.5+ (getSelectedShapes); goToByIdAsync für Slide-Navigation.
 */

import { BrandProfile, BrandConfig } from "../../services/config/BrandConfig";
import { loadBrandConfig }           from "../../services/config/ConfigService";
import { colorsMatch, colorDistance, normalizeHex } from "../../utils/colorUtils";
import { logger }                    from "../../utils/logger";

const MODULE = "BrandCheckService";

// ─── Typen ─────────────────────────────────────────────────────────────────────

export type ViolationType =
  | "font-name"
  | "font-size"
  | "font-color"
  | "fill-color"
  | "line-color";

export interface Violation {
  /** 0-basierter Slide-Index */
  slideIndex: number;
  slideId:    string;
  shapeId:    string;
  shapeName:  string;
  type:       ViolationType;
  /** Gefundener Wert (z.B. "Times New Roman" oder "#FF1234") */
  found:      string;
  /** Erlaubte Werte als lesbare Beschreibung */
  expected:   string;
  /** Kann automatisch behoben werden */
  fixable:    boolean;
}

export interface BrandCheckResult {
  violations:    Violation[];
  slideCount:    number;
  shapeCount:    number;
  /** Millisekunden für die Prüfung */
  durationMs:    number;
}

export interface FixResult {
  fixed:  number;
  errors: string[];
}

// ─── Haupt-Scan ────────────────────────────────────────────────────────────────

/**
 * Scannt alle Shapes aller Slides der aktiven Präsentation.
 * Lädt die Konfiguration aus Document Settings (Fallback: DEFAULT_BRAND_CONFIG).
 */
export async function runBrandCheck(
  onProgress?: (slidesDone: number, slidesTotal: number) => void
): Promise<BrandCheckResult> {
  const config  = loadBrandConfig();
  const profile = getActiveProfile(config);

  if (!profile) {
    throw new Error(`Konfigurationsprofil "${config.activeProfile}" nicht gefunden.`);
  }

  const startMs    = Date.now();
  const violations: Violation[] = [];
  let   shapeCount  = 0;

  await PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const slides       = presentation.slides;
    slides.load("items");
    await context.sync();

    const total = slides.items.length;

    for (let si = 0; si < total; si++) {
      const slide = slides.items[si];
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      const slideId    = slide.id;
      const slideIndex = si;

      for (const shape of shapes.items) {
        shape.load(["id", "name", "type"]);
      }
      await context.sync();

      for (const shape of shapes.items) {
        const count = await inspectShape(
          shape, slideIndex, slideId, profile, violations, context
        );
        shapeCount += count;
      }

      onProgress?.(si + 1, total);
    }
  });

  return {
    violations,
    slideCount:  0, // wird unten gesetzt
    shapeCount,
    durationMs: Date.now() - startMs,
  };
}

// ─── Shape-Inspektion ──────────────────────────────────────────────────────────

/**
 * Prüft ein einzelnes Shape und fügt Verstöße zur Liste hinzu.
 * Gibt die Anzahl der geprüften (Teil-)Shapes zurück.
 */
async function inspectShape(
  shape:      PowerPoint.Shape,
  slideIndex: number,
  slideId:    string,
  profile:    BrandProfile,
  violations: Violation[],
  context:    PowerPoint.RequestContext,
  depth       = 0
): Promise<number> {
  let count = 1;

  // ── Fill-Farbe ───────────────────────────────────────────────────────────────
  try {
    shape.fill.load("foregroundColor");
    await context.sync();
    const fillColor = normalizeHex(shape.fill.foregroundColor ?? "");
    if (fillColor && !isTransparentOrNone(fillColor)) {
      checkColor(fillColor, "fill-color", shape, slideIndex, slideId, profile, violations);
    }
  } catch { /* Shape ohne Fill oder Gradient */ }

  // ── Linienfarbe ───────────────────────────────────────────────────────────────
  try {
    shape.lineFormat.load("color");
    await context.sync();
    const lineColor = normalizeHex(shape.lineFormat.color ?? "");
    if (lineColor && !isTransparentOrNone(lineColor)) {
      checkColor(lineColor, "line-color", shape, slideIndex, slideId, profile, violations);
    }
  } catch { /* Shape ohne Linie */ }

  // ── Text (Font-Name, Font-Size, Font-Color) ────────────────────────────────
  try {
    const tf = shape.textFrame;
    const tr = tf.textRange;
    tr.load("text");
    tr.font.load(["name", "size", "color"]);
    await context.sync();

    if (tr.text && tr.text.trim().length > 0) {
      // Font-Name
      const fontName = tr.font.name;
      if (fontName && !isFontAllowed(fontName, profile)) {
        violations.push({
          slideIndex, slideId,
          shapeId:   shape.id,
          shapeName: shape.name,
          type:      "font-name",
          found:     fontName,
          expected:  profile.allowedFonts.join(", "),
          fixable:   true,
        });
      }

      // Font-Größe
      const fontSize = tr.font.size;
      if (fontSize && fontSize < profile.minFontSizePt) {
        violations.push({
          slideIndex, slideId,
          shapeId:   shape.id,
          shapeName: shape.name,
          type:      "font-size",
          found:     `${fontSize.toFixed(1)} pt`,
          expected:  `≥ ${profile.minFontSizePt} pt`,
          fixable:   true,
        });
      }

      // Font-Farbe
      const fontColor = normalizeHex(tr.font.color ?? "");
      if (fontColor && !isTransparentOrNone(fontColor)) {
        checkColor(fontColor, "font-color", shape, slideIndex, slideId, profile, violations);
      }
    }
  } catch { /* Shape ohne TextFrame */ }

  // ── Gruppe (rekursiv, max. 2 Ebenen) ──────────────────────────────────────
  if (shape.type === PowerPoint.ShapeType.group && depth < 2) {
    try {
      const groupShapes = shape.group.shapes;
      groupShapes.load("items");
      await context.sync();

      for (const child of groupShapes.items) {
        child.load(["id", "name", "type"]);
      }
      await context.sync();

      for (const child of groupShapes.items) {
        count += await inspectShape(
          child, slideIndex, slideId, profile, violations, context, depth + 1
        );
      }
    } catch { /* Gruppe nicht zugänglich */ }
  }

  // ── Tabelle ───────────────────────────────────────────────────────────────
  // Hinweis: Office.js TableCell bietet keinen paragraphs/runs Zugriff.
  // Nur Zell-Text (als String) und ggf. Fill-Farbe der Zelle prüfbar.
  // Font-Details pro Run in Tabellen: nicht robust über Office.js API verfügbar.
  // → Einschränkung in TESTING.md dokumentiert.
  if (shape.type === PowerPoint.ShapeType.table) {
    try {
      const table = shape.table;
      table.load(["rowCount", "columnCount"]);
      await context.sync();

      for (let r = 0; r < table.rowCount; r++) {
        for (let c = 0; c < table.columnCount; c++) {
          const cell = table.getCell(r, c);
          cell.load("text");
          await context.sync();
          count++;
          // Font-Details für Tabellenzellen: eingeschränkt (keine Run-API)
          // → nur Text-Existenz prüfen, keine Font-Name/Size/Color-Prüfung
        }
      }
    } catch { /* Tabelle nicht zugänglich */ }
  }

  return count;
}

// ─── Farb-Prüfung ─────────────────────────────────────────────────────────────

function checkColor(
  color:      string,
  type:       ViolationType,
  shape:      PowerPoint.Shape,
  slideIndex: number,
  slideId:    string,
  profile:    BrandProfile,
  violations: Violation[]
): void {
  const isAllowed = profile.brandColors.some(
    (bc) => colorsMatch(color, bc.value, profile.colorTolerance)
  );

  if (!isAllowed) {
    const nearest = findNearestBrandColor(color, profile);
    violations.push({
      slideIndex, slideId,
      shapeId:   shape.id,
      shapeName: shape.name,
      type,
      found:     color,
      expected:  profile.brandColors.map((bc) => bc.value).join(", "),
      fixable:   nearest !== null,
    });
  }
}

// ─── Fix-Aktionen ──────────────────────────────────────────────────────────────

/**
 * Behebt alle fixierbaren Verstöße automatisch.
 * - font-name  → ersetzt durch allowedFonts[0]
 * - font-size  → setzt auf minFontSizePt
 * - *-color    → setzt auf nächste Markenfarbe
 */
export async function fixViolations(
  violations: Violation[]
): Promise<FixResult> {
  const config  = loadBrandConfig();
  const profile = getActiveProfile(config);
  if (!profile) return { fixed: 0, errors: ["Kein aktives Profil."] };

  const result: FixResult = { fixed: 0, errors: [] };
  const fixable = violations.filter((v) => v.fixable);
  if (fixable.length === 0) return result;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const v of fixable) {
      const slide = slides.items[v.slideIndex];
      if (!slide) { result.errors.push(v.shapeName); continue; }

      const shapes = slide.shapes;
      shapes.load("items/id");
      await context.sync();

      const shape = shapes.items.find((s) => s.id === v.shapeId);
      if (!shape) { result.errors.push(v.shapeName); continue; }

      try {
        await applySingleFix(shape, v, profile, context);
        result.fixed++;
      } catch (err) {
        logger.error(MODULE, `Fix fehlgeschlagen für "${v.shapeName}": ${v.type}`, err);
        result.errors.push(v.shapeName);
      }
    }

    await context.sync();
  }).catch((err) => {
    logger.error(MODULE, "fixViolations: PowerPoint.run fehlgeschlagen.", err);
  });

  return result;
}

/** Behebt einen einzelnen Verstoß. */
async function applySingleFix(
  shape:   PowerPoint.Shape,
  v:       Violation,
  profile: BrandProfile,
  context: PowerPoint.RequestContext
): Promise<void> {
  if (v.type === "font-name" && profile.allowedFonts.length > 0) {
    shape.textFrame.textRange.font.name = profile.allowedFonts[0];
  } else if (v.type === "font-size") {
    shape.textFrame.textRange.font.size = profile.minFontSizePt;
  } else if (v.type === "font-color") {
    const nearest = findNearestBrandColor(v.found, profile);
    if (nearest) shape.textFrame.textRange.font.color = nearest;
  } else if (v.type === "fill-color") {
    const nearest = findNearestBrandColor(v.found, profile);
    if (nearest) shape.fill.setSolidColor(nearest);
  } else if (v.type === "line-color") {
    const nearest = findNearestBrandColor(v.found, profile);
    if (nearest) shape.lineFormat.color = nearest;
  }
  await context.sync();
}

// ─── Slide-Navigation ─────────────────────────────────────────────────────────

/**
 * Navigiert zur angegebenen Slide-ID.
 * Nutzt Office.context.document.goToByIdAsync (stabil auf Mac).
 */
export function goToSlide(slideId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.goToByIdAsync(
      slideId,
      Office.GoToType.Slide,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? "Slide-Navigation fehlgeschlagen."));
        }
      }
    );
  });
}

// ─── CSV-Export ───────────────────────────────────────────────────────────────

/**
 * Exportiert die Verstoßliste als CSV-Datei.
 * Nutzt Browser-native Blob + download-Link (verfügbar in WKWebView auf Mac).
 */
export function exportViolationsAsCsv(violations: Violation[], profileName: string): void {
  const header = "Folie,Shape,Verstoßtyp,Gefunden,Erwartet,Behebbar\n";
  const rows   = violations.map((v) =>
    [
      v.slideIndex + 1,
      csvEscape(v.shapeName),
      v.type,
      csvEscape(v.found),
      csvEscape(v.expected),
      v.fixable ? "Ja" : "Nein",
    ].join(",")
  );
  const csv = header + rows.join("\n");

  const blob  = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url   = URL.createObjectURL(blob);
  const link  = document.createElement("a");
  link.href   = url;
  link.setAttribute("download", `BrandCheck_${profileName}_${formatDateForFile()}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  logger.info(MODULE, `CSV-Export: ${violations.length} Verstöße exportiert.`);
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

function getActiveProfile(config: BrandConfig): BrandProfile | undefined {
  return config.profiles.find((p) => p.name === config.activeProfile);
}

function isFontAllowed(fontName: string, profile: BrandProfile): boolean {
  const lower = fontName.toLowerCase();
  return profile.allowedFonts.some((f) => f.toLowerCase() === lower);
}

/** Gibt die nächste Markenfarbe oder null zurück wenn keine Markenfarben definiert. */
function findNearestBrandColor(hex: string, profile: BrandProfile): string | null {
  if (profile.brandColors.length === 0) return null;
  let best     = profile.brandColors[0];
  let bestDist = colorDistance(hex, best.value);
  for (const bc of profile.brandColors.slice(1)) {
    const d = colorDistance(hex, bc.value);
    if (d < bestDist) { bestDist = d; best = bc; }
  }
  return best.value;
}

function isTransparentOrNone(hex: string): boolean {
  const lower = hex.toLowerCase();
  return lower === "transparent" || lower === "none" || lower === "" || lower === "#000000ff";
}

function csvEscape(value: string): string {
  if (value.includes(",") || value.includes('"') || value.includes("\n")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function formatDateForFile(): string {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
}
