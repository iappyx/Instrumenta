/**
 * SlideService.ts
 * Zentrale Operationen auf PowerPoint-Slides.
 */

import { logger } from "../../utils/logger";

const MODULE = "SlideService";

/**
 * Gibt alle Slides der Präsentation zurück.
 */
export async function getAllSlides(): Promise<PowerPoint.Slide[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    logger.debug(MODULE, `${slides.items.length} Slide(s) geladen.`);
    return slides.items;
  }).catch((err) => {
    logger.error(MODULE, "getAllSlides fehlgeschlagen.", err);
    return [];
  });
}

/**
 * Gibt die Anzahl der Slides zurück.
 */
export async function getSlideCount(): Promise<number> {
  const slides = await getAllSlides();
  return slides.length;
}

/**
 * Iteriert über alle Shapes aller Slides und ruft den Callback auf.
 * Nützlich für deck-weite Operationen.
 *
 * @param callback Wird für jedes Shape aufgerufen.
 *                 Gibt true zurück → Shape wird in die Ergebnisliste aufgenommen.
 */
export async function forEachShapeInDeck(
  callback: (shape: PowerPoint.Shape, slide: PowerPoint.Slide) => boolean | void
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      for (const shape of shapes.items) {
        shape.load(["id", "name", "type"]);
      }
      await context.sync();

      for (const shape of shapes.items) {
        callback(shape, slide);
      }
    }
  }).catch((err) => {
    logger.error(MODULE, "forEachShapeInDeck fehlgeschlagen.", err);
  });
}
