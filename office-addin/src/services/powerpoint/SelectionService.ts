/**
 * SelectionService.ts
 * Liest die aktuelle Selektion in PowerPoint (Shapes und Slides).
 * Alle Methoden geben typisierte Ergebnisse zurück und behandeln
 * den Fall "nichts selektiert" defensiv.
 */

import { logger } from "../../utils/logger";

const MODULE = "SelectionService";

export interface SelectedShape {
  id:     string;
  name:   string;
  left:   number;
  top:    number;
  width:  number;
  height: number;
  type:   string;
}

/**
 * Lädt alle selektierten Shapes mit Position und Größe.
 * Gibt ein leeres Array zurück wenn nichts selektiert oder ein Fehler auftritt.
 */
export async function getSelectedShapes(): Promise<PowerPoint.Shape[]> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    for (const shape of selection.items) {
      shape.load(["id", "name", "left", "top", "width", "height", "type", "geometricShapeType"]);
    }
    await context.sync();

    logger.debug(MODULE, `${selection.items.length} Shape(s) selektiert.`);
    return selection.items;
  }).catch((err) => {
    logger.error(MODULE, "getSelectedShapes fehlgeschlagen.", err);
    return [];
  });
}

/**
 * Gibt die aktive (erste selektierte) Slide zurück.
 * Gibt null zurück wenn keine Slide selektiert ist.
 */
export async function getActiveSlide(): Promise<PowerPoint.Slide | null> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) {
      logger.warn(MODULE, "Keine Slide selektiert.");
      return null;
    }

    const slide = slides.items[0];
    slide.load("id");
    await context.sync();

    return slide;
  }).catch((err) => {
    logger.error(MODULE, "getActiveSlide fehlgeschlagen.", err);
    return null;
  });
}

/**
 * Gibt alle selektierten Slides zurück.
 */
export async function getSelectedSlides(): Promise<PowerPoint.Slide[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();
    return slides.items;
  }).catch((err) => {
    logger.error(MODULE, "getSelectedSlides fehlgeschlagen.", err);
    return [];
  });
}
