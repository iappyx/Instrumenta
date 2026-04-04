/**
 * TextService.ts
 * Lese- und Schreiboperationen auf TextFrames und TextRuns in PowerPoint-Shapes.
 */

import { logger } from "../../utils/logger";

const MODULE = "TextService";

export interface TextRunInfo {
  text:      string;
  fontName:  string;
  fontSize:  number;
  bold:      boolean;
  italic:    boolean;
  color:     string;
}

/**
 * Liest alle TextRuns eines Shapes und gibt sie als strukturiertes Array zurück.
 * Gibt ein leeres Array zurück wenn das Shape keinen TextFrame hat.
 */
export async function getTextRuns(
  shape: PowerPoint.Shape,
  context: PowerPoint.RequestContext
): Promise<TextRunInfo[]> {
  try {
    const tf = shape.textFrame;
    tf.textRange.load(["text"]);
    tf.textRange.font.load(["name", "size", "bold", "italic", "color"]);
    await context.sync();

    // Vereinfachte Implementierung: gesamten Text als einen Run
    // Schritt 8 (Find & Replace) und Schritt 6 (Brand Check) erweitern dies
    // auf paragraph- und run-genaue Iteration.
    return [{
      text:     tf.textRange.text,
      fontName: tf.textRange.font.name,
      fontSize: tf.textRange.font.size,
      bold:     tf.textRange.font.bold,
      italic:   tf.textRange.font.italic,
      color:    tf.textRange.font.color,
    }];
  } catch (err) {
    logger.debug(MODULE, `Shape hat keinen TextFrame oder Fehler beim Lesen.`, err);
    return [];
  }
}

/**
 * Setzt die Schriftfarbe aller TextRuns eines Shapes.
 */
export function setFontColor(
  shape: PowerPoint.Shape,
  color: string
): void {
  try {
    shape.textFrame.textRange.font.color = color;
  } catch (err) {
    logger.warn(MODULE, `setFontColor fehlgeschlagen für Shape "${shape.name}".`, err);
  }
}

/**
 * Gibt den gesamten Text eines Shapes zurück (flach, ohne Formatierung).
 * Gibt null zurück wenn das Shape keinen TextFrame hat.
 */
export function getPlainText(shape: PowerPoint.Shape): string | null {
  try {
    return shape.textFrame.textRange.text;
  } catch {
    return null;
  }
}

/**
 * Setzt den gesamten Text eines Shapes (ersetzt bestehenden Text).
 * Gibt false zurück wenn das Shape keinen TextFrame hat.
 */
export function setPlainText(shape: PowerPoint.Shape, text: string): boolean {
  try {
    shape.textFrame.textRange.text = text;
    return true;
  } catch {
    return false;
  }
}
