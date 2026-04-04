/**
 * ReviewService.ts
 * Review / Annotations für das Infront Toolkit.
 *
 * Konzept:
 * - Kommentare: Textboxen (Post-It-Stil) mit Name + Zeitstempel im Text
 * - Markierungen: Gefüllte Rechtecke (Highlight-Farbe) ohne Text
 * - Shape-Namenkonvention:
 *   INFRONT_COMMENT_{timestamp}   – Kommentare
 *   INFRONT_HIGHLIGHT_{timestamp} – Markierungen
 * - Kein nativer Comments-API-Zugriff (Office.js stellt PowerPoint-Kommentare
 *   nicht zuverlässig auf Mac bereit) → Shape-basierter Ansatz
 *
 * Autor-/Zeitstempel-Format im Text: "[Vorname Nachname – dd.mm.yyyy, hh:mm]"
 * gefolgt von einer Leerzeile und dem eigentlichen Kommentartext.
 *
 * Office.js-Anforderungen:
 * - shapes.addTextBox():        1.1+
 * - shapes.addGeometricShape(): 1.4+
 * - presentation.goToSlide():   1.5+
 */

import { logger } from "../../utils/logger";

const MODULE           = "ReviewService";
const COMMENT_PREFIX   = "INFRONT_COMMENT_";
const HIGHLIGHT_PREFIX = "INFRONT_HIGHLIGHT_";

// ─── Typen ────────────────────────────────────────────────────────────────────

export interface CommentInfo {
  slideIndex:  number;
  slideId:     string;
  shapeName:   string;
  shapeId:     string;
  author:      string;
  timestamp:   string;
  commentText: string;
  preview:     string;    // Gekürzter Text für die Liste
}

export interface HighlightInfo {
  slideIndex: number;
  slideId:    string;
  shapeName:  string;
  shapeId:    string;
}

export interface AddCommentOptions {
  text:   string;
  author: string;
  /** Position auf der Slide in pt. Default: oben rechts (400/20). */
  left?:   number;
  top?:    number;
  width?:  number;
  height?: number;
  color?:  string;   // Hintergrundfarbe, default: "#FFFACD" (gelb)
}

export interface AddHighlightOptions {
  left?:   number;
  top?:    number;
  width?:  number;
  height?: number;
  color?:  string;   // default: "#FFFF00"
}

export interface ReviewStats {
  comments:   number;
  highlights: number;
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

function formatTimestamp(): string {
  return new Date().toLocaleString("de-DE", {
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit",
  });
}

function buildCommentText(author: string, commentText: string): string {
  const stamp = `[${author} – ${formatTimestamp()}]`;
  return commentText.trim() ? `${stamp}\n${commentText.trim()}` : stamp;
}

/** Parst Autor und Zeitstempel aus der ersten Zeile eines Kommentar-Texts. */
function parseCommentHeader(text: string): { author: string; timestamp: string; commentText: string } {
  const lines    = text.split("\n");
  const header   = lines[0] ?? "";
  const rest     = lines.slice(1).join("\n").trim();

  // Format: "[Name – dd.mm.yyyy, hh:mm]"
  const match = header.match(/^\[(.+?)\s*–\s*(.+?)\]$/);
  if (match) {
    return { author: match[1].trim(), timestamp: match[2].trim(), commentText: rest };
  }
  return { author: "Unbekannt", timestamp: "", commentText: text };
}

// ─── Kommentar einfügen ───────────────────────────────────────────────────────

export async function addCommentToCurrentSlide(opts: AddCommentOptions): Promise<string> {
  let shapeName = "";

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) throw new Error("Keine Folie selektiert.");
    const slide = slides.items[0];

    shapeName       = `${COMMENT_PREFIX}${Date.now()}`;
    const fullText  = buildCommentText(opts.author, opts.text);

    const box = slide.shapes.addTextBox(fullText, {
      left:   opts.left   ?? 400,
      top:    opts.top    ?? 20,
      width:  opts.width  ?? 210,
      height: opts.height ?? 80,
    });

    box.name = shapeName;

    const font = box.textFrame.textRange.font;
    font.name  = "Calibri";
    font.size  = 9;
    font.bold  = false;
    font.color = "#333333";

    box.fill.setSolidColor(opts.color ?? "#FFFACD");
    box.lineFormat.color  = "#FFA500";
    box.lineFormat.weight = 1;

    try { box.tags.add("INFRONT_REVIEW_TYPE", "comment"); } catch { /* */ }
    try { box.tags.add("INFRONT_REVIEW_AUTHOR", opts.author); } catch { /* */ }

    await context.sync();
    logger.info(MODULE, `addCommentToCurrentSlide: "${shapeName}" eingefügt.`);
  });

  return shapeName;
}

// ─── Highlight einfügen ───────────────────────────────────────────────────────

export async function addHighlightToCurrentSlide(opts: AddHighlightOptions = {}): Promise<string> {
  let shapeName = "";

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) throw new Error("Keine Folie selektiert.");
    const slide = slides.items[0];

    shapeName = `${HIGHLIGHT_PREFIX}${Date.now()}`;

    const box = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    box.name  = shapeName;
    box.load(["left", "top", "width", "height"]);
    await context.sync();

    box.left   = opts.left   ?? 100;
    box.top    = opts.top    ?? 100;
    box.width  = opts.width  ?? 200;
    box.height = opts.height ?? 80;
    box.fill.setSolidColor(opts.color ?? "#FFFF00");
    box.lineFormat.color  = "#FFA500";
    box.lineFormat.weight = 1;

    try { box.tags.add("INFRONT_REVIEW_TYPE", "highlight"); } catch { /* */ }

    await context.sync();
    logger.info(MODULE, `addHighlightToCurrentSlide: "${shapeName}" eingefügt.`);
  });

  return shapeName;
}

// ─── Alle Kommentare auflisten ────────────────────────────────────────────────

export async function findAllComments(): Promise<CommentInfo[]> {
  const found: CommentInfo[] = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (let si = 0; si < slides.items.length; si++) {
      const slide  = slides.items[si];
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      for (const shape of shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of shapes.items) {
        if (!shape.name.startsWith(COMMENT_PREFIX)) continue;

        try {
          shape.textFrame.textRange.load("text");
          await context.sync();
          const text   = shape.textFrame.textRange.text ?? "";
          const parsed = parseCommentHeader(text);

          found.push({
            slideIndex:  si,
            slideId:     slide.id,
            shapeName:   shape.name,
            shapeId:     shape.id,
            author:      parsed.author,
            timestamp:   parsed.timestamp,
            commentText: parsed.commentText,
            preview:     text.slice(0, 80) + (text.length > 80 ? "…" : ""),
          });
        } catch { /* kein TextFrame */ }
      }
    }
  });

  logger.info(MODULE, `findAllComments: ${found.length} Kommentar(e) gefunden.`);
  return found;
}

// ─── Alle Highlights auflisten ────────────────────────────────────────────────

export async function findAllHighlights(): Promise<HighlightInfo[]> {
  const found: HighlightInfo[] = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (let si = 0; si < slides.items.length; si++) {
      const slide  = slides.items[si];
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      for (const shape of shapes.items) {
        shape.load(["id", "name"]);
      }
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name.startsWith(HIGHLIGHT_PREFIX)) {
          found.push({ slideIndex: si, slideId: slide.id, shapeName: shape.name, shapeId: shape.id });
        }
      }
    }
  });

  return found;
}

// ─── Einzelnen Kommentar entfernen ────────────────────────────────────────────

export async function removeCommentByName(shapeName: string): Promise<boolean> {
  let removed = false;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name === shapeName) {
          shape.delete();
          removed = true;
          break;
        }
      }
      if (removed) break;
    }

    if (removed) await context.sync();
  });

  logger.info(MODULE, `removeCommentByName "${shapeName}": ${removed}`);
  return removed;
}

// ─── Alle Annotations entfernen ───────────────────────────────────────────────

export async function removeAllAnnotations(): Promise<ReviewStats> {
  const stats: ReviewStats = { comments: 0, highlights: 0 };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name.startsWith(COMMENT_PREFIX)) {
          shape.delete();
          stats.comments++;
        } else if (shape.name.startsWith(HIGHLIGHT_PREFIX)) {
          shape.delete();
          stats.highlights++;
        }
      }
    }

    await context.sync();
  });

  logger.info(MODULE, `removeAllAnnotations: ${stats.comments} Kommentare, ${stats.highlights} Highlights entfernt.`);
  return stats;
}

// ─── Zur Folie navigieren ──────────────────────────────────────────────────────

export async function goToSlideById(slideId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.load("id");
    }
    await context.sync();

    const target = slides.items.find((s) => s.id === slideId);
    if (target) {
      // setSelectedSlides() is on Presentation (API 1.5+), not on Slide
      context.presentation.setSelectedSlides([target.id]);
      await context.sync();
    }
  });
}

// ─── Review-Statistik ─────────────────────────────────────────────────────────

export async function getReviewStats(): Promise<ReviewStats> {
  const stats: ReviewStats = { comments: 0, highlights: 0 };

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name.startsWith(COMMENT_PREFIX))   stats.comments++;
        if (shape.name.startsWith(HIGHLIGHT_PREFIX)) stats.highlights++;
      }
    }
  });

  return stats;
}
