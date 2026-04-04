# Infront Toolkit – Testdokumentation

Manuelle Testcheckliste und API-Limitdokumentation für das Infront Toolkit.
Zielplattform: **PowerPoint für Mac (Microsoft 365)**.

---

## Testumgebungen

| Parameter | Wert |
|---|---|
| Primäre Plattform | macOS (PowerPoint für Mac) |
| Office-Version | Microsoft 365, ≥ 16.70 |
| Browser-Engine (Add-in) | WebKit (WKWebView auf Mac) |
| Node.js | ≥ 18 |
| TypeScript | ≥ 5.5 |
| Office.js Requirement Set | PowerPoint 1.5+ (angestrebt) |
| Fallback-Set | PowerPoint 1.1+ (Basis-Features) |

---

## Bekannte API-Limits und Mac-Einschränkungen

### Kategorie A – Grundsätzlich nicht umsetzbar (Office.js Mac)

| Feature | Einschränkung | Konsequenz |
|---|---|---|
| Screen-Pixel-Farbpicker | EyeDropper API nicht in WKWebView/WebKit (Stand 2026) | Fallback: Hex-Eingabe, Shape-Farbe übernehmen, Palette |
| Master vollständig ersetzen | Kein `SlideMaster`-Ersatz-API in Office.js | Fallback: Farb-/Font-Mapping deck-weit |
| Natives programm. Undo | `Application.Undo()` nicht in Office.js | Session-State-Snapshot als Fallback |
| Globale Shortcuts | Office Add-ins dürfen keine System-Shortcuts registrieren | Interne Task-Pane-Shortcuts + Ribbon-Buttons |
| Slide-Layout-Import | Layouts können nicht aus fremder Datei importiert werden | Nicht implementiert, dokumentiert in MasterImportPanel |
| Animationen programmatisch | `Shape.animations` sehr eingeschränkt | Nicht implementiert |

### Kategorie B – Umsetzbar mit Workaround

| Feature | Einschränkung | Workaround |
|---|---|---|
| Corner Radius präzise | `adjustments`-Array nur für `roundedRectangle` | Normierung via Shape-Breite/-Höhe; andere Typen werden übersprungen |
| Shape Tags persistent | API 1.5+; auf alten Versionen nicht verfügbar | Shape-Name-Konvention (`INFRONT_*`) als Primär-Identifier |
| Agenda Auto-Update | Keine `DocumentSelectionChanged`-Event-API auf Mac | Manueller Update-Button; kein automatisches Sync |
| Find & Replace Farbe | Keine Farbsuche-API | Alle Shapes iterieren, RGB-Vergleich mit konfigurierbarer Toleranz |
| Font-Farbe in Tabellen | Kein font-per-run-API für Tabellen-Zellen | Text-Farbe der gesamten Range gesetzt (grobe Annäherung) |
| Shadow-Eigenschaften | `Shape.shadow` eingeschränkt auf Mac | Try/catch; kein Fehler wenn nicht verfügbar |
| `Slide.setSelectedSlides()` | Ist auf `Presentation`, nicht `Slide` | `context.presentation.setSelectedSlides([slideId])` |
| `ShapeLineDashStyle` | Verfügbarkeit je nach API-Set | Try/catch; fällt auf Solid zurück |

### Kategorie C – Performancerisiken

| Feature | Risiko | Gegenmaßnahme |
|---|---|---|
| Brand Check (50+ Shapes) | Viele context.sync() → langsam | Batch-Load pro Slide; Fortschritts-Callback |
| Find & Replace Farbe (deck-weit) | O(n×slides×shapes) | Batch-Load pro Slide |
| Master Scan (scanDeckTheme) | Pro Shape bis zu 3 Syncs | Bewusste Entscheidung; Nutzer informiert |
| Agenda-Update (deck-weit) | context.sync() pro Shape-Schreib | Akzeptiert; wird nur bei manuellem Update ausgelöst |

---

## Modul-Übersicht

| Modul | Pfad | Zweck |
|---|---|---|
| CornerRadiusService | `src/features/cornerRadius/` | Eckenradius lesen/setzen |
| ColorPickerService | `src/features/colorPicker/` | Farben lesen/anwenden, Palette |
| BrandCheckService | `src/features/brandCheck/` | Markenprofil-Prüfung, Fixes, CSV-Export |
| FormatPainterService | `src/features/formatPainter/` | Format kopieren/anwenden, Presets |
| FindReplaceService | `src/features/findReplace/` | Text/Farbe/Font suchen+ersetzen |
| AgendaService | `src/features/agenda/` | Agenda-Shapes einfügen/aktualisieren |
| MasterImportService | `src/features/masterImport/` | Theme-Farb-/Font-Mapping (Fallback) |
| ReviewService | `src/features/review/` | Kommentare/Highlights einfügen, Navigation |
| GapEqualizerService | `src/features/gapEqualizer/` | Abstände angleichen (equal/fixed/pack) |
| RedBoxService | `src/features/redBox/` | Safe-Area-Rahmen, Foliengröße, Persistenz |
| ConfigService | `src/services/config/` | Document Settings lesen/schreiben |
| SessionState | `src/services/state/` | Undo-Stack (max 10), Format-Painter-Quelle |
| colorUtils | `src/utils/colorUtils.ts` | Hex/RGB-Konvertierung, Distanz, Normalisierung |
| geometryUtils | `src/utils/geometryUtils.ts` | Positionen/Abstände in pt, BoundingBox |
| logger | `src/utils/logger.ts` | Konsolen-Logging mit Modulpräfix |

---

## macOS Begleit-App – Integrationstests

### A. Server-Start und Health Check

- [ ] App starten → ⚡ erscheint in Menu Bar
- [ ] Setup abgeschlossen: Server startet **automatisch** beim App-Launch
- [ ] `GET http://localhost:3000/health` → `{"status":"ok","addinDistPresent":true}`
- [ ] Server stoppen (Menu Bar) → Health-Check gibt Verbindungsfehler
- [ ] Server neu starten → Health-Check antwortet wieder

### B. Brand Config Endpunkt

- [ ] `GET http://localhost:3000/config/Infront_BrandConfig.json` → gültiges JSON
- [ ] Kein User-Override: Bundle-Standard wird ausgeliefert
- [ ] User-Override unter `~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json` ablegen → Override wird ausgeliefert
- [ ] Einstellungen → Brand Config → „Auf Standard zurücksetzen" → Override gelöscht, Bundle-Standard aktiv
- [ ] Ungültiges JSON als Override → wird ignoriert, Bundle-Standard greift

### C. Remote Logging (POST /log)

- [ ] Add-in-Task Pane öffnen → Logs-Tab zeigt „Remote Logging aktiviert"-ähnliche Einträge
- [ ] `logger.info()` im Add-in → Eintrag erscheint im Logs-Tab mit `[Add-in]`-Präfix
- [ ] `logger.warn()` im Add-in → Eintrag erscheint mit Level WARN (orange)
- [ ] `logger.error()` im Add-in → Eintrag erscheint mit Level ERR (rot)
- [ ] `logger.debug()` im Add-in → wird **nicht** remote gesendet
- [ ] Server gestoppt → POST /log schlägt silent fehl, kein Fehler im Add-in

### D. Add-in Startup Integration

- [ ] Task Pane öffnen → Brand Config wird im Hintergrund vom Server geladen
- [ ] Logs-Tab: kein Eintrag für Config-Fetch-Fehler wenn macOS App nicht läuft
- [ ] macOS App läuft + Config vorhanden → Brand Check verwendet Server-Konfiguration

### E. Setup-Assistent

- [ ] Erststart: Setup-Assistent erscheint automatisch
- [ ] Schritt „Server starten" → Fortschrittsanzeige, grüne Bestätigung
- [ ] Schritt „Manifest installieren" → Datei unter `~/Library/...InfrontToolkit/Manifests/`
- [ ] Schritt „Office konfigurieren" → Katalog-Pfad sichtbar, „Pfad kopieren" funktioniert
- [ ] Schritt „Autostart" → Toggle aktiviert Login Item
- [ ] „Fertig" → Dashboard erscheint, kein Assistent mehr beim nächsten Start
- [ ] App-Neustart nach Setup → Server startet automatisch (kein Assistent)

### F. Dashboard

- [ ] Status-Tab: grüner Punkt + Port-Anzeige wenn Server läuft
- [ ] Status-Tab: Manifest-Status korrekt (installiert / nicht installiert)
- [ ] Logs-Tab: scrollt automatisch nach unten bei neuen Einträgen
- [ ] Logs-Tab: „Leeren" entfernt alle Einträge
- [ ] Einstellungen-Tab: Port-Änderung + Server-Neustart funktioniert

---

## Deployment-Voraussetzungen

```bash
npm install
npm run build        # dist/ wird erzeugt
# Oder für lokale Entwicklung:
npm run dev-server   # HTTPS-Server auf https://localhost:3000
```

Sideload in PowerPoint Mac (Entwicklung):
1. Einfügen → Add-ins → Weitere Add-ins → Add-in hochladen
2. `office-addin/manifest.xml` auswählen
3. Tab „Infront Toolkit" erscheint im Ribbon

---

## Happy-Path-Tests

### 1. Add-in laden

- [ ] Dev-Server läuft auf `https://localhost:3000`
- [ ] Zertifikat im Mac-Browser als vertrauenswürdig markiert
- [ ] `manifest.xml` erfolgreich in PowerPoint geladen (Sideload)
- [ ] Tab „Infront Toolkit" erscheint im Ribbon
- [ ] Alle 7 Gruppen sichtbar: Shapes, Format, Quality, Ausrichten, Struktur, Design, Review
- [ ] Alle deutschen Labels korrekt angezeigt

### 2. Ribbon-ExecuteFunction-Buttons

- [ ] **Gap H**: ≥3 Shapes selektiert → horizontale Abstände werden angeglichen
- [ ] **Gap V**: ≥3 Shapes selektiert → vertikale Abstände werden angeglichen
- [ ] **Red Box (Toggle)**: INFRONT_REDBOX erscheint auf aktiver Folie
- [ ] **Red Box (Toggle 2×)**: INFRONT_REDBOX wird entfernt
- [ ] **Red Box: Alle Folien**: INFRONT_REDBOX auf allen Folien eingefügt
- [ ] **Red Box entfernen**: Alle INFRONT_REDBOX-Shapes gelöscht
- [ ] **Markieren**: INFRONT_HIGHLIGHT_*-Shape erscheint auf aktiver Folie
- [ ] **Kommentare entfernen**: Alle INFRONT_COMMENT_*- und INFRONT_HIGHLIGHT_*-Shapes gelöscht

### 3. Corner Radius (Task Pane)

- [ ] Slider auf 8 pt → Rounded Rectangle angepasst
- [ ] Eingabe „0" → kein Fehler, Radius auf 0 gesetzt
- [ ] Kein Shape selektiert → Warnmeldung „Bitte ein Shape selektieren"
- [ ] Reguläres Rechteck selektiert → Meldung „0 angepasst, 1 übersprungen"
- [ ] Mehrere Shapes selektiert → alle Rounded Rectangles werden angepasst
- [ ] „Aktuellen Radius lesen" zeigt korrekten Wert

### 4. Color Picker (Task Pane)

- [ ] Hex-Eingabe `FF0000` → rote Vorschau erscheint (ohne `#` erlaubt)
- [ ] Kurzform `F00` → wird zu `#FF0000` normalisiert
- [ ] „Aus Shape" → Füllfarbe des selektierten Shapes übernommen
- [ ] Ungültiger Hex → Fehlermeldung erscheint
- [ ] „Auf Füllung anwenden" → Shape-Füllung ändert sich
- [ ] „Auf Linie anwenden" → Shape-Linie ändert sich
- [ ] „Auf Schrift anwenden" → Text-Farbe ändert sich
- [ ] Zuletzt-verwendet-Farben (max. 8) erscheinen nach erstem Anwenden
- [ ] Marken-Farben aus BrandConfig korrekt angezeigt

### 5. Brand Check (Task Pane)

- [ ] „Brand Check starten" → Fortschrittsanzeige erscheint
- [ ] Violations werden in Liste angezeigt (Typ, Folie, Wert)
- [ ] „Zur Folie" navigiert korrekt
- [ ] „Alle fixieren" → Violations werden korrigiert
- [ ] „Als CSV exportieren" → Datei-Download startet
- [ ] Präsentation ohne Violations → Meldung „Keine Verstöße gefunden"
- [ ] Profil wechseln → Liste aktualisiert sich

### 6. Format Painter+ (Task Pane)

- [ ] Shape selektieren, „Format kopieren" → Format gespeichert (Stempel-Icon aktiv)
- [ ] Anderes Shape selektieren, „Format anwenden" → Format übertragen
- [ ] Scope „Alle Folien, gleicher Layout-Typ" → deck-weite Anwendung
- [ ] Preset speichern → Name erscheint in Preset-Liste
- [ ] Preset laden → Format wird eingefügt
- [ ] Preset löschen → aus Liste entfernt

### 7. Find & Replace (Task Pane)

- [ ] **Text-Tab**: Suchen + Ersetzen → Treffer werden ersetzt, Anzahl gemeldet
- [ ] **Text-Tab**: Groß-/Kleinschreibung-Flag funktioniert
- [ ] **Text-Tab**: Regex-Flag funktioniert (z.B. `\d+` findet Zahlen)
- [ ] **Farbe-Tab**: Hex-Eingabe Quell-/Ziel-Farbe → Shapes werden ersetzt
- [ ] **Font-Tab**: Schriftart von Arial auf Calibri → deck-weit ersetzt
- [ ] Leeres Suchfeld → Buttons deaktiviert
- [ ] Keine Treffer → Meldung „0 Ersetzungen"

### 8. Agenda Wizard (Task Pane)

- [ ] 3 Abschnitte eingeben → Shapes INFRONT_AGENDA_ITEM_01-03 auf aktiver Folie
- [ ] „Auf aktiver Folie" → nur aktuelle Folie hat Agenda
- [ ] „Alle aktualisieren" → alle INFRONT_AGENDA_*-Shapes werden aktualisiert
- [ ] „Alle entfernen" → alle Agenda-Shapes gelöscht
- [ ] „Agenda-Folien finden" → zeigt Folien-Indizes mit Agenda-Shapes

### 9. Master / Theme Import (Task Pane)

- [ ] API-Einschränkungs-Warnung sichtbar
- [ ] **Deck-Scan**: Farben + Fonts erscheinen nach Scan
- [ ] Farb-Swatch klicken → wird als Quell-Farbe in Mapping übernommen
- [ ] Font klicken → wird als Quell-Font in Mapping übernommen
- [ ] **Preset laden**: Mappings werden befüllt
- [ ] Farb-Mapping eingeben + anwenden → Shapes werden gefärbt
- [ ] Font-Mapping eingeben + anwenden → Schriftarten werden ersetzt
- [ ] Toleranz 0 = exakt; Toleranz 20 = toleranter Match

### 10. Review / Annotationen (Task Pane)

- [ ] Kommentar einfügen → INFRONT_COMMENT_*-Shape auf aktiver Folie (gelbe Textbox)
- [ ] Zeitstempel korrekt formatiert: `[Name – dd.mm.yyyy, hh:mm]`
- [ ] Highlight einfügen → INFRONT_HIGHLIGHT_*-Rechteck auf aktiver Folie
- [ ] Farb-Auswahl für Kommentar/Highlight funktioniert
- [ ] Statistik-Banner zeigt korrekte Anzahl
- [ ] **Meine Kommentare**: Liste zeigt alle INFRONT_COMMENT_*-Shapes
- [ ] „Zur Folie" → navigiert zu korrekter Folie
- [ ] Einzelnen Kommentar löschen → aus Liste und Folie entfernt
- [ ] „Alle entfernen" → alle Annotations-Shapes gelöscht

### 11. Gap Equalizer (Task Pane)

- [ ] **Equal, Horizontal**: 4 Shapes → mittlere 2 gleichmäßig verteilt
- [ ] **Fixed, 12 pt**: 3 Shapes → Abstand exakt 12 pt zwischen allen
- [ ] **Pack, Vertikal**: Shapes ohne Lücke aufgestapelt
- [ ] „Vorschau" zeigt berechneten Abstand ohne zu ändern
- [ ] Negative Vorschau → Überlappungs-Warnung erscheint
- [ ] < 3 Shapes (equal) → Warnmeldung
- [ ] Schnell-Buttons setzen Modus korrekt

### 12. Red Box (Task Pane)

- [ ] Abstände 20/20/20/20 pt → Box korrekt positioniert (720×540 Folie)
- [ ] Abstände ändern → CSS-Vorschau aktualisiert sich sofort
- [ ] „Verknüpft"-Toggle: Wert in einem Feld → alle vier Felder gleich
- [ ] Farb-Preset → Vorschau aktualisiert sich
- [ ] Linientyp Gestrichelt → Box hat gestrichelte Linie
- [ ] Status-Badge: grün wenn Box vorhanden, grau wenn nicht
- [ ] „Aktualisieren" → Margins und Farbe auf bestehender Box neu gesetzt
- [ ] „Alle Folien" → Box auf allen Folien ohne Box eingefügt
- [ ] „Alle entfernen" → alle INFRONT_REDBOX-Shapes gelöscht
- [ ] Einstellungen werden nach Neustart der Task Pane geladen (Document.Settings)

---

## Edge-Case-Tests

| Testfall | Erwartetes Verhalten |
|---|---|
| Keine Shapes selektiert | Alle relevanten Features zeigen Warnmeldung, kein Absturz |
| Leere Präsentation (0 Folien) | Kein Absturz; ggf. Meldung „Keine Folien gefunden" |
| Gruppenselektierung | Gruppe wird iteriert (max. 2 Ebenen tief); einzelne Shapes werden verarbeitet |
| Tabellen-Shape selektiert | Corner Radius: übersprungen; Brand Check: Text-Farbe geprüft, kein Font-per-Run |
| Shapes mit 0 Breite/Höhe | Gap-Equalizer rechnet korrekt, kein NaN |
| Red Box: Margins > Foliengröße | Fehler „Breite/Höhe ≤ 0" wird gemeldet |
| Hex-Eingabe mit #-Prefix | normalizeHex() entfernt `#` korrekt |
| Hex-Eingabe ohne #-Prefix | normalizeHex() ergänzt `#` korrekt |
| Ungültiger Hex | normalizeHex() gibt null zurück; UI zeigt Fehlermeldung |
| Font-Replace: Schriftart nicht vorhanden | Kein Absturz; Meldung „0 ersetzt" |
| context.sync() Fehler (Network) | Catch-Block gibt Fehlermeldung an NotificationBar weiter |
| PowerPoint.run Timeout | Catch-Block gibt Fehlermeldung aus |
| Document.Settings nicht verfügbar | getSetting() gibt defaultValue zurück, kein Absturz |
| 50+ Shapes auf einer Folie | Brand Check < 5s; kein Timeout |
| Emoji und Sonderzeichen im Text | Find & Replace findet und ersetzt korrekt |
| Regex mit `/g` Flag | lastIndex wird vor jedem test() zurückgesetzt (Bug-Prävention) |
| Gradient Fill | FormatPainter: `unsupported: true`; kein Absturz |
| Shadow-Eigenschaften | try/catch verhindert Absturz auf Mac |
| Tags API nicht verfügbar (API < 1.5) | try/catch verhindert Absturz; Shape-Name-Fallback aktiv |

---

## Regressionstest nach Änderungen

Vor jedem Release die folgenden Smoke-Tests durchführen:

1. **Ribbon lädt**: Tab erscheint, alle Gruppen sichtbar, keine JS-Fehler in DevTools
2. **Corner Radius**: Rounded Rectangle, Wert 8 → Radius wird gesetzt
3. **Red Box Toggle**: Einmal ein, einmal aus auf aktueller Folie
4. **Gap H**: 4 Shapes → gleichmäßige Verteilung
5. **Kommentar einfügen**: Textbox auf aktiver Folie mit Zeitstempel
6. **Brand Check**: Scan startet ohne Absturz

---

## Deployment-Checkliste (DMG-Release)

### Vor dem Build

- [ ] `VERSION`-Datei aktualisiert (z.B. `1.1.0`)
- [ ] `office-addin/package.json`: `"version"` angepasst
- [ ] `LocalServerService.swift`: Version in `buildHealthResponse()` aktualisiert
- [ ] `CHANGELOG.md` aktualisiert
- [ ] Alle Tests grün (Happy Path + Edge Cases oben)

### Build

```bash
# Vollständiger Release-Build
SKIP_SIGNING=1 sh scripts/build-release.sh   # Entwicklung
# oder:
CODESIGN_IDENTITY="..." sh scripts/build-release.sh  # Produktion

# Einzeln:
sh scripts/build-addin.sh         # → office-addin/dist/
sh scripts/copy-dist-to-app.sh    # → mac-app/Resources/addin-dist/
sh scripts/build-mac-app.sh       # → .build/mac-app/.../InfrontToolkit.app
sh scripts/package-dmg.sh         # → .build/dist/InfrontToolkit-<version>.dmg
```

- [ ] `build-release.sh` ohne Fehler abgeschlossen
- [ ] `office-addin/dist/taskpane.js` vorhanden
- [ ] `office-addin/dist/commands.js` vorhanden
- [ ] `mac-app/Resources/addin-dist/` nicht leer
- [ ] `InfrontToolkit.app` unter `.build/mac-app/Build/Products/Release/`
- [ ] DMG unter `.build/dist/InfrontToolkit-<version>.dmg`

### DMG testen

```bash
VERSION=$(cat VERSION | tr -d '[:space:]')
open ".build/dist/InfrontToolkit-${VERSION}.dmg"
```

- [ ] DMG öffnet sich, App + /Applications-Symlink sichtbar
- [ ] App in /Applications kopieren und starten
- [ ] ⚡ Menu Bar Icon erscheint
- [ ] Setup-Assistent startet (Erststart)
- [ ] Server lässt sich starten
- [ ] Manifest wird installiert
- [ ] PowerPoint: Ribbon-Tab erscheint nach Office-Catalog-Setup
- [ ] Health-Check: `curl http://localhost:3000/health` → `{"status":"ok",...}`

### Produktion (Code Signing + Notarization)

```bash
spctl -a -vvv -t open ".build/dist/InfrontToolkit-${VERSION}.dmg"
# → accepted  source=Notarized Developer ID
```

- [ ] Notarization: `spctl` zeigt `accepted`
- [ ] Kein Gatekeeper-Dialog beim ersten Start
- [ ] `codesign --verify --deep --strict` ohne Fehler

### Release veröffentlichen

- [ ] Git-Tag anlegen: `git tag -a "v${VERSION}" -m "Release v${VERSION}"`
- [ ] Tag pushen: `git push origin "v${VERSION}"`
- [ ] GitHub Release mit DMG-Anhang erstellen
- [ ] `INSTALL_MAC.md` an Nutzer kommuniziert

---

## Bekannte Nicht-Implementierte Features (Scope)

Diese Features wurden bewusst ausgelassen oder werden in einer späteren Version umgesetzt:

| Feature | Begründung |
|---|---|
| SlideMaster vollständig ersetzen | Technisch nicht möglich (API-Grenze) |
| Animations-Editor | Außerhalb Scope v1.0 |
| Slide-Sortierer / Folien verschieben | Kein API in Office.js |
| PDF-Export konfigurieren | Kein API in Office.js |
| Inhaltsverzeichnis auto-generieren | Spätere Version |
| Mehrsprachigkeit (EN/DE Toggle) | Spätere Version |
