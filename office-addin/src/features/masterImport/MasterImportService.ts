/**
 * MasterImportService.ts
 * Master / Theme importieren für das Infront Toolkit.
 *
 * ── API-Grenzen auf Mac / Office.js ──────────────────────────────────────────
 * Office.js stellt KEINE API bereit um:
 * - SlideMaster vollständig zu ersetzen
 * - SlideMaster aus einer anderen .pptx-Datei zu importieren
 * - Theme-XML direkt zu lesen oder zu schreiben
 * - SlideLayouts hinzuzufügen oder umzubenennen
 *
 * Dokumentiert in TESTING.md (Kategorie A: Grundsätzlich nicht umsetzbar).
 *
 * ── Machbarer Fallback ────────────────────────────────────────────────────────
 * 1. Theme-Farb-Mapping: Suche + Ersatz bestehender Farben über FindReplaceService
 * 2. Theme-Font-Mapping:  Suche + Ersatz bestehender Fonts über FindReplaceService
 * 3. Preset-Farb-Sets:   Infront-Standard-Palette in einem Schritt anwenden
 * 4. Deck-Diagnose:       Welche Farben und Fonts werden aktuell verwendet?
 *
 * Die Theme-Farb/-Font-Ersetzungen delegieren an FindReplaceService,
 * da dieser bereits robuste Office.js-Batch-Logik implementiert.
 */

import { logger }       from "../../utils/logger";
import { normalizeHex } from "../../utils/colorUtils";
import {
  replaceColor,
  replaceFont,
  collectFonts,
  ColorQuery,
  FontReplaceOptions,
} from "../findReplace/FindReplaceService";

const MODULE = "MasterImportService";

// ─── Typen ────────────────────────────────────────────────────────────────────

export interface ThemeColorMapping {
  from: string;   // Hex-Farbe (Quelle)
  to:   string;   // Hex-Farbe (Ziel)
  label?: string; // Optionaler Label (z.B. "Primärfarbe")
}

export interface ThemeFontMapping {
  from: string;   // Schriftart (Quelle)
  to:   string;   // Schriftart (Ziel)
}

export interface ThemeApplyOptions {
  targetFill:       boolean;
  targetLine:       boolean;
  targetFont:       boolean;
  colorTolerance:   number;
  keepFontSize:     boolean;
  keepFontFormatting: boolean;
}

export interface ThemeApplyResult {
  colorsReplaced: number;
  fontsReplaced:  number;
  errors:         string[];
}

export interface DeckThemeSnapshot {
  colors: string[];   // Alle gefundenen Fill-/Line-/Font-Farben (dedupliziert)
  fonts:  string[];   // Alle gefundenen Schriftarten
}

// ─── Infront-Standard-Presets ─────────────────────────────────────────────────

export interface ThemePreset {
  name:         string;
  colorMappings: ThemeColorMapping[];
  fontMappings:  ThemeFontMapping[];
}

export const INFRONT_PRESETS: ThemePreset[] = [
  {
    name: "Infront Standard → Präsentation",
    colorMappings: [
      { from: "#0070C0", to: "#003366", label: "Primärblau → Infront Navy" },
      { from: "#00B0F0", to: "#003366", label: "Hellblau → Infront Navy" },
      { from: "#FF0000", to: "#CC0000", label: "Standard-Rot → Infront Rot" },
    ],
    fontMappings: [
      { from: "Arial", to: "Calibri" },
      { from: "Times New Roman", to: "Calibri" },
    ],
  },
  {
    name: "Office-Standard → Infront",
    colorMappings: [
      { from: "#4472C4", to: "#003366", label: "Office-Blau → Infront Navy" },
      { from: "#ED7D31", to: "#CC0000", label: "Office-Orange → Infront Rot" },
      { from: "#A9D18E", to: "#AAAAAA", label: "Office-Grün → Grau" },
    ],
    fontMappings: [
      { from: "Calibri Light", to: "Calibri" },
    ],
  },
];

// ─── Default-Optionen ─────────────────────────────────────────────────────────

export const DEFAULT_THEME_OPTIONS: ThemeApplyOptions = {
  targetFill:         true,
  targetLine:         true,
  targetFont:         true,
  colorTolerance:     10,
  keepFontSize:       true,
  keepFontFormatting: true,
};

// ─── Theme anwenden ───────────────────────────────────────────────────────────

/**
 * Wendet ein komplettes Theme-Mapping (Farben + Fonts) auf das gesamte Deck an.
 * Delegiert an FindReplaceService für robuste Batch-Verarbeitung.
 */
export async function applyThemeMapping(
  colorMappings: ThemeColorMapping[],
  fontMappings:  ThemeFontMapping[],
  options:       ThemeApplyOptions
): Promise<ThemeApplyResult> {
  const result: ThemeApplyResult = { colorsReplaced: 0, fontsReplaced: 0, errors: [] };

  // Farb-Ersetzungen
  for (const mapping of colorMappings) {
    const fromHex = normalizeHex(mapping.from);
    const toHex   = normalizeHex(mapping.to);
    if (!fromHex || !toHex) {
      result.errors.push(`Ungültige Farbe: ${mapping.from} → ${mapping.to}`);
      continue;
    }

    const colorQuery: ColorQuery = {
      searchHex:  fromHex,
      replaceHex: toHex,
      tolerance:  options.colorTolerance,
      targetFill: options.targetFill,
      targetLine: options.targetLine,
      targetFont: options.targetFont,
    };

    try {
      const r = await replaceColor(colorQuery, "allSlides");
      result.colorsReplaced += r.replaced;
      if (r.errors.length > 0) {
        result.errors.push(...r.errors.map((e) => `Farbe ${mapping.from}: ${e}`));
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      logger.error(MODULE, `replaceColor ${mapping.from} → ${mapping.to}: ${msg}`, err);
      result.errors.push(`Farbe ${mapping.from}: ${msg}`);
    }
  }

  // Font-Ersetzungen
  const fontOptions: FontReplaceOptions = {
    keepSize:       options.keepFontSize,
    keepFormatting: options.keepFontFormatting,
  };

  for (const mapping of fontMappings) {
    if (!mapping.from.trim() || !mapping.to.trim()) continue;
    try {
      const r = await replaceFont(mapping.from, mapping.to, fontOptions, "allSlides");
      result.fontsReplaced += r.replaced;
      if (r.errors.length > 0) {
        result.errors.push(...r.errors.map((e) => `Font ${mapping.from}: ${e}`));
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      logger.error(MODULE, `replaceFont ${mapping.from} → ${mapping.to}: ${msg}`, err);
      result.errors.push(`Font ${mapping.from}: ${msg}`);
    }
  }

  logger.info(MODULE, `applyThemeMapping: ${result.colorsReplaced} Farbe(n), ${result.fontsReplaced} Font(s) ersetzt.`);
  return result;
}

// ─── Deck-Diagnose ─────────────────────────────────────────────────────────────

/**
 * Scannt das Deck und sammelt alle verwendeten Farben und Fonts.
 * Farben: Fill (Solid), Line, Font.
 * Fonts: über collectFonts() aus FindReplaceService.
 */
export async function scanDeckTheme(): Promise<DeckThemeSnapshot> {
  const colorSet = new Set<string>();
  const fonts    = await collectFonts();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        // Fill-Farbe
        try {
          shape.fill.load("foregroundColor");
          await context.sync();
          const hex = normalizeHex(shape.fill.foregroundColor ?? "");
          if (hex) colorSet.add(hex);
        } catch { /* kein Fill */ }

        // Linie-Farbe
        try {
          shape.lineFormat.load("color");
          await context.sync();
          const hex = normalizeHex(shape.lineFormat.color ?? "");
          if (hex) colorSet.add(hex);
        } catch { /* keine Linie */ }

        // Font-Farbe
        try {
          shape.textFrame.textRange.font.load("color");
          await context.sync();
          const hex = normalizeHex(shape.textFrame.textRange.font.color ?? "");
          if (hex) colorSet.add(hex);
        } catch { /* kein Text */ }
      }
    }
  });

  const colors = Array.from(colorSet).sort();
  logger.info(MODULE, `scanDeckTheme: ${colors.length} Farben, ${fonts.length} Fonts gefunden.`);
  return { colors, fonts };
}
