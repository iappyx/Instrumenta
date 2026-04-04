/**
 * ShapeService.ts
 * Zentrale Operationen auf PowerPoint-Shapes.
 * Alle Methoden arbeiten innerhalb eines PowerPoint.run()-Kontexts.
 */

import { logger } from "../../utils/logger";

const MODULE = "ShapeService";

/** Shape-Typen, die Text-Frames unterstützen. */
const TEXT_SUPPORTING_TYPES = new Set([
  PowerPoint.ShapeType.textBox,
  PowerPoint.ShapeType.placeholder,
  PowerPoint.ShapeType.freeform,
]);

/**
 * Prüft ob ein Shape einen TextFrame besitzt (defensiv).
 * Vorsicht: shape.textFrame kann auf bestimmten Shape-Typen eine Exception werfen.
 */
export function supportsText(shape: PowerPoint.Shape): boolean {
  try {
    return TEXT_SUPPORTING_TYPES.has(shape.type) || shape.textFrame !== null;
  } catch {
    return false;
  }
}

/**
 * Sucht ein Shape nach Name auf einer Slide.
 * Gibt null zurück wenn nicht gefunden.
 */
export async function findShapeByName(
  slide: PowerPoint.Slide,
  name: string,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape | null> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  const found = shapes.items.find((s) => s.name === name);
  if (!found) {
    logger.debug(MODULE, `Shape "${name}" nicht gefunden.`);
    return null;
  }
  return found;
}

/**
 * Sucht alle Shapes mit einem bestimmten Name-Präfix auf einer Slide.
 */
export async function findShapesByPrefix(
  slide: PowerPoint.Slide,
  prefix: string,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape[]> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  return shapes.items.filter((s) => s.name.startsWith(prefix));
}

/**
 * Löscht ein Shape nach Name von einer Slide.
 * Gibt true zurück wenn gelöscht, false wenn nicht gefunden.
 */
export async function deleteShapeByName(
  slide: PowerPoint.Slide,
  name: string,
  context: PowerPoint.RequestContext
): Promise<boolean> {
  const shape = await findShapeByName(slide, name, context);
  if (!shape) return false;
  shape.delete();
  await context.sync();
  logger.info(MODULE, `Shape "${name}" gelöscht.`);
  return true;
}

/**
 * Gibt alle Shapes eines Decks zurück, die mit einem Präfix benannt sind.
 * Nützlich für globale Operationen (Red Box entfernen, Kommentare entfernen).
 */
export async function findAllShapesByPrefixInDeck(
  prefix: string
): Promise<Array<{ slideId: string; shape: PowerPoint.Shape }>> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const results: Array<{ slideId: string; shape: PowerPoint.Shape }> = [];

    for (const slide of slides.items) {
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name.startsWith(prefix)) {
          results.push({ slideId: slide.id, shape });
        }
      }
    }

    return results;
  }).catch((err) => {
    logger.error(MODULE, "findAllShapesByPrefixInDeck fehlgeschlagen.", err);
    return [];
  });
}
