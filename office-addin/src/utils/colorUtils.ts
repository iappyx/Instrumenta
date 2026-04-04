/**
 * colorUtils.ts
 * Hilfsfunktionen für Farb-Parsing, -Konvertierung und -Vergleich.
 */

export interface RGB {
  r: number;
  g: number;
  b: number;
}

/**
 * Konvertiert einen Hex-String (#RRGGBB) in ein RGB-Objekt.
 * Gibt null zurück wenn das Format ungültig ist.
 */
export function hexToRgb(hex: string): RGB | null {
  const cleaned = hex.replace(/^#/, "");
  if (!/^[0-9A-Fa-f]{6}$/.test(cleaned)) return null;
  return {
    r: parseInt(cleaned.slice(0, 2), 16),
    g: parseInt(cleaned.slice(2, 4), 16),
    b: parseInt(cleaned.slice(4, 6), 16),
  };
}

/**
 * Konvertiert ein RGB-Objekt in einen Hex-String (#RRGGBB).
 */
export function rgbToHex(rgb: RGB): string {
  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`.toUpperCase();
}

function toHex(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");
}

/**
 * Berechnet den euklidischen Farbabstand zwischen zwei Hex-Farben im RGB-Raum.
 * Wertebereich: 0 (identisch) bis ~441 (schwarz/weiß).
 */
export function colorDistance(hex1: string, hex2: string): number {
  const a = hexToRgb(hex1);
  const b = hexToRgb(hex2);
  if (!a || !b) return 999;
  return Math.sqrt(
    Math.pow(a.r - b.r, 2) +
    Math.pow(a.g - b.g, 2) +
    Math.pow(a.b - b.b, 2)
  );
}

/**
 * Prüft ob zwei Farben innerhalb einer Toleranz liegen.
 * Toleranzwert 0–30 entspricht den Infront-Brand-Check-Einstellungen.
 * Hinweis: Toleranz 30 im RGB-Raum entspricht einem Abstand von ~52 im euklidischen Sinne.
 */
export function colorsMatch(hex1: string, hex2: string, tolerance: number = 0): boolean {
  const dist = colorDistance(hex1, hex2);
  // Skalierung: tolerance 0-30 → RGB-Abstand-Schwelle 0-52
  const threshold = tolerance * (52 / 30);
  return dist <= threshold;
}

/**
 * Normalisiert eine Farbe auf #RRGGBB-Format.
 * Akzeptiert: #RGB, #RRGGBB, RRGGBB (ohne #).
 * Gibt null zurück bei ungültigem Input.
 */
export function normalizeHex(raw: string): string | null {
  const cleaned = raw.trim().replace(/^#/, "");
  if (/^[0-9A-Fa-f]{6}$/.test(cleaned)) return `#${cleaned.toUpperCase()}`;
  if (/^[0-9A-Fa-f]{3}$/.test(cleaned)) {
    const [r, g, b] = cleaned.split("");
    return `#${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }
  return null;
}

/**
 * Berechnet die relative Luminanz einer Farbe (für Kontrastverhältnis).
 */
export function luminance(hex: string): number {
  const rgb = hexToRgb(hex);
  if (!rgb) return 0;
  const toLinear = (c: number): number => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * toLinear(rgb.r) + 0.7152 * toLinear(rgb.g) + 0.0722 * toLinear(rgb.b);
}
