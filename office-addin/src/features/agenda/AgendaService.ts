/**
 * AgendaService.ts
 * Agenda-Wizard für das Infront Toolkit.
 *
 * Konzept:
 * - Bis zu 8 Sektionen mit Name + Foliennummer
 * - Eine separate Textbox pro Sektion auf den Agenda-Slides
 * - Shape-Namenkonvention: INFRONT_AGENDA_ITEM_01 – INFRONT_AGENDA_ITEM_08
 * - Konfiguration (Sektionen, Formatierung) in Document.Settings
 * - Shape Tags als optionales Supplement (PowerPoint API 1.5+, try/catch)
 * - Kein eventbasiertes Auto-Update auf Mac → manueller Update-Button
 *
 * Shape-Layout auf Standard-Folie (720 × 540 pt):
 * - Jede Sektion: Textbox mit name + optionaler Foliennummer
 * - Stapelweise von oben nach unten, links ausgerichtet
 *
 * Office.js-Anforderungen:
 * - Shape.name: 1.1+  |  shapes.addTextBox(): 1.1+
 * - Shape.tags: 1.5+ (optional, try/catch)
 * - Document.settings: 1.1+
 */

import { logger }       from "../../utils/logger";
import { getSetting, setSetting, flushSettings } from "../../services/config/ConfigService";

const MODULE      = "AgendaService";
const CONFIG_KEY  = "INFRONT_AGENDA_CONFIG";
const SHAPE_PREFIX = "INFRONT_AGENDA_ITEM_";

// ─── Typen ─────────────────────────────────────────────────────────────────────

export interface AgendaSection {
  name:        string;
  slideNumber: number;  // 1-basiert
}

export interface AgendaFormat {
  color:    string;
  fontSize: number;
  bold:     boolean;
  italic:   boolean;
}

export interface AgendaConfig {
  sections:       AgendaSection[];
  activeFormat:   AgendaFormat;
  inactiveFormat: AgendaFormat;
  showPageNumbers: boolean;
  /** Aktive Sektion (0-basiert), -1 = keine */
  activeIndex:    number;
  /** Layout: linke Position der Textboxen in pt */
  left:   number;
  /** Layout: obere Startposition in pt */
  top:    number;
  /** Layout: Breite der Textboxen in pt */
  width:  number;
  /** Layout: Höhe pro Textbox in pt */
  itemHeight: number;
  /** Layout: vertikaler Abstand zwischen Textboxen in pt */
  gap:    number;
}

export interface AgendaUpdateResult {
  updated:  number;
  errors:   string[];
}

// ─── Default-Konfiguration ────────────────────────────────────────────────────

export const DEFAULT_AGENDA_CONFIG: AgendaConfig = {
  sections: Array.from({ length: 8 }, (_, i) => ({ name: "", slideNumber: 0 })),
  activeFormat:   { color: "#003366", fontSize: 12, bold: true,  italic: false },
  inactiveFormat: { color: "#AAAAAA", fontSize: 11, bold: false, italic: false },
  showPageNumbers: true,
  activeIndex:    0,
  left:   24,
  top:    80,
  width:  280,
  itemHeight: 32,
  gap:    4,
};

// ─── Konfiguration laden / speichern ──────────────────────────────────────────

export function loadAgendaConfig(): AgendaConfig {
  const saved = getSetting<AgendaConfig | null>(CONFIG_KEY, null);
  if (saved) return { ...DEFAULT_AGENDA_CONFIG, ...saved };
  return { ...DEFAULT_AGENDA_CONFIG };
}

export async function saveAgendaConfig(config: AgendaConfig): Promise<void> {
  setSetting(CONFIG_KEY, config);
  await flushSettings();
  logger.info(MODULE, "Agenda-Konfiguration gespeichert.");
}

// ─── Shape-Name Helper ────────────────────────────────────────────────────────

function itemShapeName(index: number): string {
  return `${SHAPE_PREFIX}${String(index + 1).padStart(2, "0")}`;
}

function parseItemIndex(shapeName: string): number | null {
  if (!shapeName.startsWith(SHAPE_PREFIX)) return null;
  const n = parseInt(shapeName.slice(SHAPE_PREFIX.length), 10);
  return isNaN(n) ? null : n - 1;
}

// ─── Agenda-Text formatieren ──────────────────────────────────────────────────

function buildItemText(section: AgendaSection, isActive: boolean, showPage: boolean): string {
  const marker = isActive ? "▶  " : "     ";
  const page   = showPage && section.slideNumber > 0 ? `  ${section.slideNumber}` : "";
  return `${marker}${section.name}${page}`;
}

// ─── Shape einfügen ───────────────────────────────────────────────────────────

/**
 * Fügt Agenda-Shapes auf der aktuellen Slide ein.
 * Leere Sektionen werden übersprungen.
 * Bestehende Agenda-Shapes auf dieser Slide werden zuerst gelöscht.
 */
export async function insertAgendaOnCurrentSlide(config: AgendaConfig): Promise<AgendaUpdateResult> {
  const result: AgendaUpdateResult = { updated: 0, errors: [] };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) {
      logger.warn(MODULE, "Keine Slide selektiert.");
      return;
    }

    const slide = slides.items[0];

    // Bestehende Agenda-Shapes auf dieser Slide entfernen
    await removeAgendaFromSlide(slide, context);

    const activeSections = config.sections.filter((s) => s.name.trim().length > 0);
    if (activeSections.length === 0) {
      logger.warn(MODULE, "Keine Sektionen konfiguriert.");
      return;
    }

    let topOffset = config.top;

    for (let i = 0; i < activeSections.length; i++) {
      const section  = activeSections[i];
      const isActive = i === config.activeIndex;
      const fmt      = isActive ? config.activeFormat : config.inactiveFormat;
      const text     = buildItemText(section, isActive, config.showPageNumbers);

      try {
        const box = slide.shapes.addTextBox(text, {
          left:   config.left,
          top:    topOffset,
          width:  config.width,
          height: config.itemHeight,
        });

        box.name = itemShapeName(i);

        const font = box.textFrame.textRange.font;
        font.name   = "Calibri";
        font.size   = fmt.fontSize;
        font.bold   = fmt.bold;
        font.italic = fmt.italic;
        font.color  = fmt.color;

        box.fill.setNoFill();

        // Shape Tags als optionales Supplement
        try {
          box.tags.add("INFRONT_AGENDA_INDEX", String(i));
          box.tags.add("INFRONT_AGENDA_ACTIVE", String(isActive));
        } catch { /* Tags nicht unterstützt – kein Problem */ }

        topOffset += config.itemHeight + config.gap;
        result.updated++;
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        logger.error(MODULE, `Sektion ${i + 1} einfügen fehlgeschlagen: ${msg}`, err);
        result.errors.push(`Sektion ${i + 1}`);
      }
    }

    await context.sync();
    logger.info(MODULE, `${result.updated} Agenda-Shapes eingefügt.`);
  });

  return result;
}

// ─── Alle Agenda-Shapes aktualisieren ─────────────────────────────────────────

/**
 * Aktualisiert alle Agenda-Shapes im gesamten Deck.
 * Findet Shapes nach Namenspräfix und wendet Formatierung neu an.
 */
export async function updateAllAgendaShapes(config: AgendaConfig): Promise<AgendaUpdateResult> {
  const result: AgendaUpdateResult = { updated: 0, errors: [] };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const activeSections = config.sections.filter((s) => s.name.trim().length > 0);

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        const idx = parseItemIndex(shape.name);
        if (idx === null) continue;

        const section = activeSections[idx];
        if (!section) continue;

        const isActive = idx === config.activeIndex;
        const fmt      = isActive ? config.activeFormat : config.inactiveFormat;
        const text     = buildItemText(section, isActive, config.showPageNumbers);

        try {
          shape.textFrame.textRange.text = text;
          const font = shape.textFrame.textRange.font;
          font.size   = fmt.fontSize;
          font.bold   = fmt.bold;
          font.italic = fmt.italic;
          font.color  = fmt.color;

          // Tag aktualisieren
          try { shape.tags.add("INFRONT_AGENDA_ACTIVE", String(isActive)); } catch { /* */ }

          result.updated++;
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          logger.error(MODULE, `Update Shape "${shape.name}": ${msg}`, err);
          result.errors.push(shape.name);
        }
      }
    }

    await context.sync();
    logger.info(MODULE, `updateAllAgendaShapes: ${result.updated} Shape(s) aktualisiert.`);
  });

  return result;
}

// ─── Agenda-Shapes entfernen ──────────────────────────────────────────────────

/**
 * Entfernt alle Agenda-Shapes aus dem gesamten Deck.
 */
export async function removeAllAgendaShapes(): Promise<number> {
  let removed = 0;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const n = await removeAgendaFromSlide(slide, context);
      removed += n;
    }

    await context.sync();
  });

  logger.info(MODULE, `removeAllAgendaShapes: ${removed} Shape(s) entfernt.`);
  return removed;
}

/** Entfernt Agenda-Shapes von einer einzelnen Slide. Gibt Anzahl zurück. */
async function removeAgendaFromSlide(
  slide:   PowerPoint.Slide,
  context: PowerPoint.RequestContext
): Promise<number> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  let count = 0;
  for (const shape of shapes.items) {
    if (shape.name.startsWith(SHAPE_PREFIX)) {
      shape.delete();
      count++;
    }
  }
  return count;
}

// ─── Agenda-Shapes im Deck suchen ─────────────────────────────────────────────

/**
 * Gibt alle Slides zurück, die Agenda-Shapes enthalten (mit Shape-Count).
 */
export async function findAgendaSlides(): Promise<Array<{ slideIndex: number; count: number }>> {
  const found: Array<{ slideIndex: number; count: number }> = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (let si = 0; si < slides.items.length; si++) {
      const slide  = slides.items[si];
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      const count = shapes.items.filter((s) => s.name.startsWith(SHAPE_PREFIX)).length;
      if (count > 0) found.push({ slideIndex: si, count });
    }
  });

  return found;
}
