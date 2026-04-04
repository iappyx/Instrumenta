/**
 * FindReplaceService.ts
 * Suchen & Ersetzen für Text, Farben und Schriftarten im gesamten Deck.
 *
 * ── Text-Ersatz ──────────────────────────────────────────────────────────────
 * Implementiert über textRange.text (Gesamttext-Ersatz).
 *
 * WICHTIG – API-Einschränkung:
 * `textRange.text = newValue` ersetzt den kompletten Text eines Shapes.
 * Intra-Shape-Formatierungen (einzelne fette Wörter, gemischte Schriftgrößen
 * innerhalb eines Textblocks) gehen dabei verloren.
 * Eine vollständig format-bewahrende Ersetzung wäre nur über Paragraph/Run-Level
 * möglich – Office.js PowerPoint stellt hierfür keine ausreichende API bereit.
 * Dokumentiert in TESTING.md (Kategorie B).
 *
 * ── Farb-Ersatz ──────────────────────────────────────────────────────────────
 * Vergleich mit konfigurierbarer Toleranz (0–30) via colorUtils.colorsMatch().
 * Anwendbar auf: Fill, Line, Font getrennt oder kombiniert.
 *
 * ── Font-Ersatz ──────────────────────────────────────────────────────────────
 * Sammelt alle im Deck verwendeten Fonts, ermöglicht Font-zu-Font-Ersatz.
 * Optionen: Schriftgröße beibehalten, Formatierung (Bold/Italic) beibehalten.
 *
 * Office.js-Anforderungen:
 * - Shape.textFrame / textRange: 1.1+
 * - Shape.fill / lineFormat: 1.1+ / 1.4+
 * - getSelectedSlides(): 1.5+
 */

import { logger }       from "../../utils/logger";
import { colorsMatch, normalizeHex } from "../../utils/colorUtils";
import { pushUndo, createSnapshot }  from "../../services/state/SessionState";

const MODULE = "FindReplaceService";

// ─── Gemeinsame Typen ──────────────────────────────────────────────────────────

export type SearchScope = "currentSlide" | "allSlides" | "selectedSlides";

export interface ReplaceResult {
  replaced:   number;
  previewed:  number;
  errors:     string[];
}

// ─── Text-Typen ───────────────────────────────────────────────────────────────

export interface TextQuery {
  search:        string;
  replacement:   string;
  caseSensitive: boolean;
  wholeWord:     boolean;
}

export interface TextMatch {
  slideIndex: number;
  slideId:    string;
  shapeName:  string;
  shapeId:    string;
  preview:    string;
  count:      number;
}

// ─── Farb-Typen ───────────────────────────────────────────────────────────────

export interface ColorQuery {
  searchHex:     string;
  replaceHex:    string;
  tolerance:     number;
  targetFill:    boolean;
  targetLine:    boolean;
  targetFont:    boolean;
}

export interface ColorMatch {
  slideIndex: number;
  shapeName:  string;
  shapeId:    string;
  target:     "fill" | "line" | "font";
  foundColor: string;
}

// ─── Font-Typen ───────────────────────────────────────────────────────────────

export interface FontReplaceOptions {
  keepSize:        boolean;
  keepFormatting:  boolean;
}

// ─── Scope Helper ─────────────────────────────────────────────────────────────

async function getSlidesForScope(
  scope:   SearchScope,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Slide[]> {
  if (scope === "currentSlide" || scope === "selectedSlides") {
    const sel = context.presentation.getSelectedSlides();
    sel.load("items");
    await context.sync();
    if (scope === "currentSlide") {
      return sel.items.length > 0 ? [sel.items[0]] : [];
    }
    return sel.items;
  }

  // allSlides
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TEXT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Erstellt eine Regex für den Suchtext.
 * Sonderzeichen werden escaped. `wholeWord` nutzt \b auf ASCII und
 * alternativ einen Leerzeichen/Satzzeichen-Workaround für Umlaute.
 */
function buildTextRegex(query: TextQuery): RegExp {
  const escaped = query.search.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern  = query.wholeWord ? `(?<![\\wäöüÄÖÜß])${escaped}(?![\\wäöüÄÖÜß])` : escaped;
  const flags    = query.caseSensitive ? "g" : "gi";
  return new RegExp(pattern, flags);
}

/**
 * Vorschau: Findet alle Textvorkommen ohne zu ersetzen.
 */
export async function previewText(
  query: TextQuery,
  scope: SearchScope
): Promise<TextMatch[]> {
  if (!query.search) return [];
  const regex   = buildTextRegex(query);
  const matches: TextMatch[] = [];

  await PowerPoint.run(async (context) => {
    const slides = await getSlidesForScope(scope, context);

    for (let si = 0; si < slides.length; si++) {
      const slide = slides[si];
      slide.load("id");
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of slide.shapes.items) {
        try {
          shape.textFrame.textRange.load("text");
          await context.sync();
          const text = shape.textFrame.textRange.text;
          if (!text) continue;

          const found = text.match(regex);
          if (found && found.length > 0) {
            matches.push({
              slideIndex: si,
              slideId:    slide.id,
              shapeName:  shape.name,
              shapeId:    shape.id,
              preview:    text.slice(0, 60) + (text.length > 60 ? "…" : ""),
              count:      found.length,
            });
          }
        } catch { /* kein TextFrame */ }
      }
    }
  });

  return matches;
}

/**
 * Ersetzt alle Textvorkommen im Scope.
 *
 * HINWEIS: Verwendet textRange.text-Zuweisung → intra-Shape-Formatierung
 * (gemischte Schriftgrößen, partielles Bold etc.) geht verloren.
 */
export async function replaceText(
  query: TextQuery,
  scope: SearchScope
): Promise<ReplaceResult> {
  if (!query.search) return { replaced: 0, previewed: 0, errors: [] };

  const regex   = buildTextRegex(query);
  const result: ReplaceResult = { replaced: 0, previewed: 0, errors: [] };
  const snapshots = [];

  await PowerPoint.run(async (context) => {
    const slides = await getSlidesForScope(scope, context);

    for (let si = 0; si < slides.length; si++) {
      const slide = slides[si];
      slide.load("id");
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of slide.shapes.items) {
        try {
          shape.textFrame.textRange.load("text");
          await context.sync();

          const original = shape.textFrame.textRange.text;
          if (!original) continue;

          // Regex zurücksetzen (lastIndex)
          regex.lastIndex = 0;
          if (!regex.test(original)) continue;
          regex.lastIndex = 0;

          const replaced = original.replace(regex, query.replacement);
          if (replaced === original) continue;

          snapshots.push(createSnapshot(slide.id, shape.id, shape.name, { text: original }));
          shape.textFrame.textRange.text = replaced;
          result.replaced++;
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          logger.warn(MODULE, `replaceText Shape "${shape.name}": ${msg}`);
          result.errors.push(shape.name);
        }
      }
    }

    await context.sync();
  });

  if (snapshots.length > 0) {
    pushUndo({ featureName: "Text Suchen & Ersetzen", timestamp: Date.now(), snapshots });
  }

  logger.info(MODULE, `replaceText: ${result.replaced} Shape(s) geändert.`);
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// FARBE
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Vorschau: Findet alle Shapes mit der Suchfarbe (innerhalb Toleranz).
 */
export async function previewColor(
  query: ColorQuery,
  scope: SearchScope
): Promise<ColorMatch[]> {
  const searchHex = normalizeHex(query.searchHex);
  if (!searchHex) return [];

  const matches: ColorMatch[] = [];

  await PowerPoint.run(async (context) => {
    const slides = await getSlidesForScope(scope, context);

    for (let si = 0; si < slides.length; si++) {
      const slide = slides[si];
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of slide.shapes.items) {
        if (query.targetFill) {
          try {
            shape.fill.load("foregroundColor");
            await context.sync();
            const hex = normalizeHex(shape.fill.foregroundColor ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              matches.push({ slideIndex: si, shapeName: shape.name, shapeId: shape.id, target: "fill", foundColor: hex });
            }
          } catch { /* kein Fill */ }
        }

        if (query.targetLine) {
          try {
            shape.lineFormat.load("color");
            await context.sync();
            const hex = normalizeHex(shape.lineFormat.color ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              matches.push({ slideIndex: si, shapeName: shape.name, shapeId: shape.id, target: "line", foundColor: hex });
            }
          } catch { /* keine Linie */ }
        }

        if (query.targetFont) {
          try {
            shape.textFrame.textRange.font.load("color");
            await context.sync();
            const hex = normalizeHex(shape.textFrame.textRange.font.color ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              matches.push({ slideIndex: si, shapeName: shape.name, shapeId: shape.id, target: "font", foundColor: hex });
            }
          } catch { /* kein Text */ }
        }
      }
    }
  });

  return matches;
}

/**
 * Ersetzt alle Farben im Scope.
 */
export async function replaceColor(
  query: ColorQuery,
  scope: SearchScope
): Promise<ReplaceResult> {
  const searchHex  = normalizeHex(query.searchHex);
  const replaceHex = normalizeHex(query.replaceHex);
  if (!searchHex || !replaceHex) {
    return { replaced: 0, previewed: 0, errors: ["Ungültige Hex-Farbe."] };
  }

  const result: ReplaceResult = { replaced: 0, previewed: 0, errors: [] };
  const snapshots = [];

  await PowerPoint.run(async (context) => {
    const slides = await getSlidesForScope(scope, context);

    for (let si = 0; si < slides.length; si++) {
      const slide = slides[si];
      slide.load("id");
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of slide.shapes.items) {
        let changed = false;

        if (query.targetFill) {
          try {
            shape.fill.load("foregroundColor");
            await context.sync();
            const hex = normalizeHex(shape.fill.foregroundColor ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              snapshots.push(createSnapshot(slide.id, shape.id, shape.name, { fill: hex }));
              shape.fill.setSolidColor(replaceHex);
              changed = true;
            }
          } catch { /* kein Fill */ }
        }

        if (query.targetLine) {
          try {
            shape.lineFormat.load("color");
            await context.sync();
            const hex = normalizeHex(shape.lineFormat.color ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              snapshots.push(createSnapshot(slide.id, shape.id, shape.name, { line: hex }));
              shape.lineFormat.color = replaceHex;
              changed = true;
            }
          } catch { /* keine Linie */ }
        }

        if (query.targetFont) {
          try {
            shape.textFrame.textRange.font.load("color");
            await context.sync();
            const hex = normalizeHex(shape.textFrame.textRange.font.color ?? "");
            if (hex && colorsMatch(hex, searchHex, query.tolerance)) {
              snapshots.push(createSnapshot(slide.id, shape.id, shape.name, { fontColor: hex }));
              shape.textFrame.textRange.font.color = replaceHex;
              changed = true;
            }
          } catch { /* kein Text */ }
        }

        if (changed) result.replaced++;
      }
    }

    await context.sync();
  });

  if (snapshots.length > 0) {
    pushUndo({ featureName: "Farbe Suchen & Ersetzen", timestamp: Date.now(), snapshots });
  }

  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// FONT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Sammelt alle im Deck verwendeten Schriftarten.
 * Iteriert alle Slides und Shapes mit TextFrame.
 */
export async function collectFonts(): Promise<string[]> {
  const fonts = new Set<string>();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        try {
          shape.textFrame.textRange.font.load("name");
          await context.sync();
          const name = shape.textFrame.textRange.font.name;
          if (name && name.trim()) fonts.add(name.trim());
        } catch { /* kein TextFrame */ }
      }
    }
  });

  return Array.from(fonts).sort((a, b) => a.localeCompare(b));
}

/**
 * Ersetzt einen Schriftart durch eine andere im gesamten Scope.
 */
export async function replaceFont(
  fromFont: string,
  toFont:   string,
  options:  FontReplaceOptions,
  scope:    SearchScope
): Promise<ReplaceResult> {
  if (!fromFont || !toFont) {
    return { replaced: 0, previewed: 0, errors: ["Bitte Quell- und Ziel-Schriftart wählen."] };
  }

  const result: ReplaceResult = { replaced: 0, previewed: 0, errors: [] };
  const snapshots = [];
  const fromLower  = fromFont.toLowerCase();

  await PowerPoint.run(async (context) => {
    const slides = await getSlidesForScope(scope, context);

    for (let si = 0; si < slides.length; si++) {
      const slide = slides[si];
      slide.load("id");
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of slide.shapes.items) {
        try {
          const font = shape.textFrame.textRange.font;
          font.load(["name", "size", "bold", "italic"]);
          await context.sync();

          if (!font.name || font.name.toLowerCase() !== fromLower) continue;

          snapshots.push(createSnapshot(slide.id, shape.id, shape.name, {
            fontName: font.name,
            fontSize: font.size,
            bold:     font.bold,
            italic:   font.italic,
          }));

          font.name = toFont;
          if (!options.keepSize && font.size)         { /* Größe beibehalten: nichts tun */ }
          if (!options.keepFormatting) {
            font.bold   = false;
            font.italic = false;
          }

          result.replaced++;
        } catch { /* kein TextFrame */ }
      }
    }

    await context.sync();
  });

  if (snapshots.length > 0) {
    pushUndo({ featureName: "Font Suchen & Ersetzen", timestamp: Date.now(), snapshots });
  }

  logger.info(MODULE, `replaceFont "${fromFont}" → "${toFont}": ${result.replaced} Shape(s)`);
  return result;
}
