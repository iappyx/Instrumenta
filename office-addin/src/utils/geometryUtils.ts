/**
 * geometryUtils.ts
 * Hilfsfunktionen für Positions- und Größenberechnungen in Punkten (pt).
 * PowerPoint-Koordinaten sind in Punkten (pt): 1 pt = 1/72 Zoll.
 */

export interface Rect {
  left:   number;
  top:    number;
  width:  number;
  height: number;
}

/** Gibt den rechten Rand eines Shapes zurück. */
export function right(r: Rect): number {
  return r.left + r.width;
}

/** Gibt den unteren Rand eines Shapes zurück. */
export function bottom(r: Rect): number {
  return r.top + r.height;
}

/** Gibt den horizontalen Mittelpunkt zurück. */
export function centerX(r: Rect): number {
  return r.left + r.width / 2;
}

/** Gibt den vertikalen Mittelpunkt zurück. */
export function centerY(r: Rect): number {
  return r.top + r.height / 2;
}

/**
 * Konvertiert Pixel in Punkte (bei 96 DPI).
 * 1 px bei 96 DPI = 0.75 pt.
 */
export function pxToPt(px: number): number {
  return px * 0.75;
}

/**
 * Konvertiert Punkte in Pixel (bei 96 DPI).
 */
export function ptToPx(pt: number): number {
  return pt / 0.75;
}

/**
 * Berechnet den gleichmäßigen Abstand zwischen N Shapes (Kantenabstand).
 * Voraussetzung: shapes sind nach position (left oder top) sortiert.
 * @param totalSpan   Gesamtspanne (right(last) - left(first) oder bottom(last) - top(first))
 * @param totalSize   Summe aller Shape-Breiten oder -Höhen
 * @param count       Anzahl der Shapes
 */
export function equalGap(totalSpan: number, totalSize: number, count: number): number {
  if (count <= 1) return 0;
  return (totalSpan - totalSize) / (count - 1);
}

/**
 * Prüft ob zwei Shapes sich überlappen (horizontal).
 */
export function overlapsHorizontally(a: Rect, b: Rect): boolean {
  return a.left < right(b) && right(a) > b.left;
}

/**
 * Prüft ob zwei Shapes sich überlappen (vertikal).
 */
export function overlapsVertically(a: Rect, b: Rect): boolean {
  return a.top < bottom(b) && bottom(a) > b.top;
}

/**
 * Gibt den Begrenzungsrahmen einer Liste von Rects zurück.
 */
export function boundingBox(rects: Rect[]): Rect {
  if (rects.length === 0) return { left: 0, top: 0, width: 0, height: 0 };
  const minLeft = Math.min(...rects.map((r) => r.left));
  const minTop  = Math.min(...rects.map((r) => r.top));
  const maxRight  = Math.max(...rects.map((r) => right(r)));
  const maxBottom = Math.max(...rects.map((r) => bottom(r)));
  return {
    left:   minLeft,
    top:    minTop,
    width:  maxRight  - minLeft,
    height: maxBottom - minTop,
  };
}
