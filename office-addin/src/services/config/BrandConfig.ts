/**
 * BrandConfig.ts
 * Typdefinitionen für die Brand-Compliance-Konfiguration.
 * Konfigurationsdatei: config/Infront_BrandConfig.json
 */

export interface BrandColorEntry {
  /** Anzeigename der Markenfarbe */
  name:  string;
  /** Hex-Wert, z.B. "#003366" */
  value: string;
}

export interface BrandProfile {
  /** Profilname, z.B. "Default" oder "Strict" */
  name: string;

  /** Erlaubte Schriftarten (Groß-/Kleinschreibung wird ignoriert beim Vergleich) */
  allowedFonts: string[];

  /** Erlaubte Markenfarben */
  brandColors: BrandColorEntry[];

  /**
   * Farbtoleranz für den Farbvergleich.
   * Bereich: 0 (exakter Match) bis 30 (großzügig).
   * Entspricht einem euklidischen RGB-Abstand von ca. 0–52.
   */
  colorTolerance: number;

  /** Minimale erlaubte Schriftgröße in Punkten */
  minFontSizePt: number;
}

export interface BrandConfig {
  /** Schema-Version für zukünftige Migrations-Unterstützung */
  version: string;

  /** Aktives Profil (muss einem Profilnamen in `profiles` entsprechen) */
  activeProfile: string;

  /** Alle verfügbaren Konfigurationsprofile */
  profiles: BrandProfile[];
}

/** Standard-Konfiguration (entspricht Infront_BrandConfig.ini) */
export const DEFAULT_BRAND_CONFIG: BrandConfig = {
  version: "1.0",
  activeProfile: "Default",
  profiles: [
    {
      name:           "Default",
      allowedFonts:   ["Calibri", "Arial", "Helvetica Neue"],
      brandColors: [
        { name: "Infront Navy",  value: "#003366" },
        { name: "Infront Red",   value: "#FF0000" },
        { name: "Weiß",          value: "#FFFFFF" },
        { name: "Schwarz",       value: "#000000" },
        { name: "Hellgrau",      value: "#CCCCCC" },
      ],
      colorTolerance: 10,
      minFontSizePt:  8,
    },
    {
      name:           "Strict",
      allowedFonts:   ["Calibri"],
      brandColors: [
        { name: "Infront Navy",  value: "#003366" },
        { name: "Weiß",          value: "#FFFFFF" },
        { name: "Schwarz",       value: "#000000" },
      ],
      colorTolerance: 0,
      minFontSizePt:  10,
    },
  ],
};
