# Changelog – Infront Toolkit

Alle wesentlichen Änderungen werden in dieser Datei dokumentiert.
Format: [Semantic Versioning](https://semver.org) · [Keep a Changelog](https://keepachangelog.com)

---

## [1.1.0-rc4] – 2026-04-04

DMG Build und Release-Struktur implementiert (Schritt 7).

### Neu

- **`VERSION`**: Einzelne Versionsnummer-Datei als Single Source of Truth für alle Build-Skripte
- **`scripts/package-dmg.sh`**: Vollständig implementiert:
  - Liest Version aus `VERSION`-Datei
  - Methode 1: `create-dmg` (brew) für hübsches DMG mit App-Fenster
  - Methode 2: `hdiutil` (macOS built-in, immer verfügbar) als Fallback
  - Code Signing via `CODESIGN_IDENTITY` Umgebungsvariable
  - Notarization via `xcrun notarytool` + `NOTARIZATION_PROFILE`
  - `SKIP_SIGNING=1` für Development / CI ohne Zertifikate
  - Ausgabe: `.build/dist/InfrontToolkit-<version>.dmg`
- **`scripts/build-release.sh`**: Release-Orchestrierung — führt alle 4 Schritte in Reihe aus (Add-in → copy-dist → macOS App → DMG)
- **`docs/RELEASE.md`**: Vollständige Release-Dokumentation:
  - Version aktualisieren (alle Stellen)
  - Code Signing + Notarization einrichten (einmalig)
  - Development vs. Production Build
  - DMG testen + Notarization verifizieren
  - GitHub Release anlegen (inkl. `gh`-CLI Snippet)
  - Post-Release Checkliste
  - Fehlerbehebung (xcodebuild, notarization, Gatekeeper)

---

## [1.1.0-rc3] – 2026-04-04

End-user Installation Flow implementiert (Schritt 6).

### Neu – Dokumentation

- **`docs/INSTALL_MAC.md`**: Vollständige Benutzeranleitung für die macOS-Installation:
  - Voraussetzungen (macOS 13+, Microsoft 365)
  - DMG-Installation + Gatekeeper-Hinweis
  - Setup-Assistent walkthrough (alle 6 Schritte)
  - Normaler Betrieb + Brand Config anpassen
  - Fehlerbehebung (5 Szenarien)
  - Deinstallation + Konfigurationspfade

### Geändert – macOS App

- **`Models/AppState.swift`** (fix): `manifestInstalledPath` wird in `init()` aus dem Dateisystem wiederhergestellt — nach App-Neustart wird das Manifest korrekt als installiert erkannt
- **`Views/SetupWizardView.swift`** (verbessert):
  - Schritt „Office konfigurieren": Katalog-Ordner-Pfad jetzt immer sichtbar (nicht conditional)
  - Neue Buttons: „Ordner in Finder öffnen" + „Pfad kopieren" (direkt in Zwischenablage)
  - Anweisungstext überarbeitet: Schritt 6 erwähnt explizit dass der Ribbon-Tab erscheint

---

## [1.1.0-rc2] – 2026-04-04

Lokale Assets, Konfiguration und Startlogik implementiert (Schritt 5).

### Neu – macOS App

- **`Models/AppState.swift`** (erweitert): `startIfNeeded()` – startet den Server automatisch beim App-Launch wenn Setup abgeschlossen; wird im `init()` via `Task { @MainActor }` ausgelöst (kein Dependency auf Window-Lifecycle)

### Neu – Add-in (office-addin/)

- **`src/taskpane/index.tsx`** (erweitert): macOS-Integration beim Startup:
  - `enableRemoteLogging(macAppUrl)` — aktiviert Remote Logging zur Begleit-App
  - `fetchBrandConfigFromServer()` + `saveBrandConfig()` im Hintergrund — lädt aktuelle Brand Config vom Server, fällt bei Fehler still zurück

### Geändert – macOS App

- **`Views/SettingsView.swift`** (erweitert): "Brand Config"-Sektion hinzugefügt:
  - Zeigt ob User-Override oder Bundle-Standard aktiv ist
  - Öffnet User-Config-Ordner im Finder (erstellt ihn ggf.)
  - "Auf Standard zurücksetzen" (löscht User-Override) wenn aktiv

---

## [1.1.0-rc1] – 2026-04-04

Office.js ↔ native macOS Integration implementiert (Schritt 4).

### Neu – macOS App

- **`Services/ConfigSyncService.swift`**: Brand-Konfiguration verwalten; Priorität: User-Override → Bundle-Default → Eingebettetes Fallback; `currentConfigData()`, `saveUserConfig()`, `resetToDefault()`
- **`Services/LocalServerService.swift`** (erweitert): Vollständiges Routing hinzugefügt:
  - `GET /health` → `{"status":"ok","version":"1.1.0","addinDistPresent":bool}`
  - `GET /config/Infront_BrandConfig.json` → über `ConfigSyncService.currentConfigData()`
  - `POST /log` → Log-Eintrag vom Add-in; dispatcht an `logHandler` auf `@MainActor`
  - `setLogHandler()` für Dependency Injection
- **`Models/AppState.swift`** (erweitert): `setupAddinLogHandler()` nach Server-Start; leitet Add-in-Logs mit `[Add-in]`-Präfix in den Logs-Tab

### Neu – Add-in (office-addin/)

- **`src/utils/logger.ts`** (erweitert): Remote Logging hinzugefügt:
  - `enableRemoteLogging(baseUrl)`: Aktiviert fire-and-forget POST /log
  - `disableRemoteLogging()`: Deaktiviert Remote Logging
  - Debug-Logs werden nicht remote gesendet; Fehler silent ignoriert
- **`src/services/config/ConfigService.ts`** (erweitert): macOS Integration hinzugefügt:
  - `getMacAppBaseUrl()`: Gibt `http://localhost:<port>` zurück
  - `fetchBrandConfigFromServer()`: Optionaler HTTP-Fetch mit 3 s Timeout; fällt bei Fehler auf `null` zurück

### Geändert

- **`scripts/copy-dist-to-app.sh`**: Vollständig implementiert; kopiert nun auch `office-addin/config/` → `addin-dist/config/` (Brand-Config für Bundle-Default)

### Neu – Dokumentation

- **`docs/INTEGRATION.md`**: HTTP-Endpunkte, Kommunikationsmodell, Sicherheitsbetrachtung, Konfigurationspfade

---

## [1.1.0-beta] – 2026-04-04

Native macOS Begleit-App Grundgerüst implementiert (Schritt 3).

### Neu – mac-app/ Grundgerüst

- **`InfrontToolkitApp.swift`**: Menu Bar App (`@main`, `MenuBarExtra`, `Window` Scene), kein Dock-Icon (`LSUIElement`)
- **`Models/AppState.swift`**: `@MainActor ObservableObject` mit `ServerStatus`, `SetupStep`, `LogEntry`; Server-Control-Methoden; UserDefaults-Persistenz
- **`Models/ServerConfig.swift`**: Port, Host, `addinDistURL` (Bundle-Ressource), `isAddinDistPresent` Prüfung
- **`Services/LocalServerService.swift`**: Minimaler HTTP-Server via `NWListener` (Network.framework); nur loopback (127.0.0.1); Path-Traversal-Prävention; MIME-Type-Erkennung; CORS-Header
- **`Services/ManifestService.swift`**: Manifest aus Template generieren (`{{PORT}}` ersetzen); in `~/Library/Application Support/InfrontToolkit/Manifests/` installieren
- **`Services/AppAutoStartService.swift`**: `SMAppService.mainApp` (macOS 13+) für Login Item Registrierung
- **`Views/StatusMenuView.swift`**: Menu Bar Dropdown (Status, Dashboard öffnen, Server starten/stoppen, Beenden)
- **`Views/SetupWizardView.swift`**: 6-Schritte Erststart-Assistent (Willkommen → Server → Manifest → Office → Autostart → Fertig)
- **`Views/MainWindowView.swift`**: Dashboard (Status-Tab + Logs-Tab) + Umschaltung auf Setup-Wizard bei Erststart
- **`Views/SettingsView.swift`**: Port, Autostart-Toggle, Manifest-Verwaltung
- **`Resources/manifest-template.xml`**: Vollständiges Office-Manifest mit `{{PORT}}` Platzhalter (HTTP statt HTTPS für localhost)
- **`Info.plist`**: `LSUIElement=YES` (kein Dock-Icon), `LSMinimumSystemVersion=13.0`
- **`InfrontToolkit.entitlements`**: App Sandbox, `network.server`, `network.client`, `files.user-selected.read-write`
- **`Assets.xcassets`**: Platzhalter App Icon Set (Icons noch zu ergänzen)
- **`InfrontToolkit.xcodeproj/project.pbxproj`**: Vollständige Xcode Projektdatei (alle Sources, Resources, Build Settings)
- **`project.yml`**: XcodeGen Konfiguration als wartbare Alternative zu project.pbxproj
- **`mac-app/README.md`**: Build-Anleitung, Projektstruktur, bekannte Einschränkungen

### Geändert

- `scripts/build-mac-app.sh`: Vollständig implementiert (war Stub); kopiert addin-dist/, prüft xcodegen, ruft xcodebuild auf

### Architekturentscheidung: HTTP statt HTTPS für localhost

Die macOS App serviert den Add-in `dist/` über plain HTTP (nicht HTTPS).
Office für Mac erlaubt `http://localhost:PORT` für Shared Folder Catalogs.
Vorteil: Kein selbstsigniertes Zertifikat nötig, kein Keychain-Trust-Problem.
Nachteil: Nur für localhost (loopback) – nicht für externe Deployments geeignet.

---

## [1.1.0-alpha] – 2026-04-04

Beginn der macOS-Migration: Umbau zur deploybaren macOS-Lösung mit nativer Begleit-App und DMG-Distribution.

### Neu – Repo-Struktur (Schritt 2)

- **`office-addin/`**: Alle Add-in-Quelldateien zentralisiert (`src/`, `config/`, `manifest.xml`, Build-Konfiguration)
- **`mac-app/`**: Verzeichnis für native Swift/SwiftUI Begleit-App (Grundgerüst folgt in Schritt 3)
- **`scripts/`**: Build- und Packaging-Skripte
  - `setup-dev.sh`: Entwicklungsumgebung einrichten
  - `build-addin.sh`: Production Build des Add-ins
  - `build-mac-app.sh`: macOS App bauen (Stub, Schritt 3)
  - `copy-dist-to-app.sh`: dist/ in App Bundle kopieren (Stub, Schritt 4)
  - `package-dmg.sh`: DMG erstellen (Stub, Schritt 7)
- **`docs/`**: Architekturdokumentation
  - `ARCHITECTURE.md`: Architekturentscheidungen, Komponentenverantwortung, Installationsfluss
- **`legacy/`**: Archivierter VBA-Code und alte Installer (nicht aktiv)
  - `legacy/vba-source/`: Classes, Forms, Modules, CustomUI, Installers
  - `legacy/bin/`: Alte Binaries (.ppam, .dmg, .exe)
  - `legacy/version-notes/`: Versionsnotizen v0.8–v1.69
- **`office-addin/assets/icons/`**: Platzhalter-Icons (16/32/80 px, Infront Navy) erstellt

### Geändert

- `office-addin/package.json`: `package`-Script ergänzt (`npm run build && copy-dist-to-app.sh`)
- `README.md`: Komplett auf neue Zwei-Komponenten-Architektur aktualisiert

### Architekturentscheidung

Der lokale HTTP-Server-Ansatz (Ansatz C) wurde als Zielarchitektur gewählt:
- Die native macOS App bundelt den gebauten Add-in `dist/` und serviert ihn lokal
- `manifest.xml` SourceLocation zeigt auf `http://localhost:3000` (Office erlaubt localhost)
- Keine Cloud-Abhängigkeit für die Add-in-Assets
- Nutzer-Onboarding über Setup-Assistent in der macOS App

Bekannte Einschränkungen dieser Architektur dokumentiert in: `docs/ARCHITECTURE.md`

---

## [1.0.0] – 2026-04-04

Erste vollständige Version des Infront Toolkits als Office.js Add-in für PowerPoint auf Mac.
Ersetzt das VBA-basierte Instrumenta-Add-in vollständig.

### Neu – Infrastruktur (Schritte 2+3)

- **Architektur**: React 18 + TypeScript 5.5 (strict) + Webpack 5 + Fluent UI v8
- **manifest.xml**: Klassisches XML-Manifest (stabiler als Unified Manifest auf Mac)
  - GUID: `3f7d5a9e-2b4c-4f8a-9d1e-6c0b8a3e5f7d`
  - 7 Ribbon-Gruppen: Shapes, Format, Quality, Ausrichten, Struktur, Design, Review
  - 18 Ribbon-Buttons (Mix aus ShowTaskpane + ExecuteFunction)
  - Deutsche Labels, Screentips, Supertips
  - `keytip`-Attribute für alle Buttons vorbereitet
- **Routing**: URL-Parameter `?view=` ohne externe Router-Abhängigkeit
- **office-js**: Vom CDN geladen (kein Bundling), externe `office-js`-Deklaration in Webpack
- **HTTPS Dev-Server**: `https://localhost:3000` mit Hot Reload

### Neu – Feature-Implementierungen

#### Schritt 4: Corner Radius (`src/features/cornerRadius/`)
- `CornerRadiusService.ts`: Normierung `adjustments[0]` auf `roundedRectangle`-Shapes
- Formel: `normalized = ptValue / min(width, height) × 2`, geclampt auf `[0, 1]`
- Nicht-kompatible Shape-Typen werden übersprungen (kein Absturz)
- Liest aktuellen Radius (`readCurrentRadiusPx`) und setzt neu (`applyCornerRadius`)

#### Schritt 5: Color Picker (`src/features/colorPicker/`)
- `ColorPickerService.ts`: Fill, Line, Font-Farben lesen und setzen
- Kein EyeDropper (WKWebView-Limitation) → Hex-Eingabe + Shape-Farbe übernehmen + Palette
- `recentColors[]`: Session-persistent, max. 8 Einträge, FIFO
- Marken-Farben aus `BrandConfig` integriert

#### Schritt 6: Brand Check (`src/features/brandCheck/`)
- `BrandCheckService.ts`: Prüft Font-Name, Font-Größe, Font-Farbe, Fill-Farbe, Line-Farbe
- Gruppen-Rekursion bis 2 Ebenen tief
- Fortschritts-Callback für UI-Updates
- `fixViolations()`: Batch-Fix aller oder ausgewählter Violations
- `exportViolationsAsCsv()`: `Blob + URL.createObjectURL()` – funktioniert in WKWebView
- Profil-Unterstützung: Default + Strict aus `Infront_BrandConfig.json`
- Navigation zur betroffenen Folie via `goToSlide(slideId)`

#### Schritt 7: Format Painter+ (`src/features/formatPainter/`)
- `FormatPainterService.ts`: Erfasst Fill, Line, Font-Eigenschaften
- Gradient-Fill erkannt als `unsupported: true` (kein Absturz)
- Shadow via try/catch (Mac-limitiert)
- Scopes: `selection`, `slideByType`, `deckByType`
- Presets: speichern/laden/löschen via `Document.Settings`

#### Schritt 8: Find & Replace (`src/features/findReplace/`)
- `FindReplaceService.ts`: Text, Farbe (mit Toleranz), Schriftart
- Regex-Bug-Prävention: `regex.lastIndex = 0` vor jedem `test()`
- `textRange.text`-Zuweisung verliert Intra-Shape-Formatierung (dokumentiert)
- Toleranz-basierter Farb-Vergleich via `colorUtils.colorsMatch()`

#### Schritt 9: Agenda Wizard (`src/features/agenda/`)
- `AgendaService.ts`: Shapes `INFRONT_AGENDA_ITEM_01` bis `INFRONT_AGENDA_ITEM_08`
- Separate Textbox pro Abschnitt (nicht Paragraphen)
- Shape-Tags als optionales Supplement (try/catch, API 1.5+)
- Kein Auto-Update (kein zuverlässiges `DocumentSelectionChanged`-Event auf Mac)

#### Schritt 10: Master / Theme Import (`src/features/masterImport/`)
- `MasterImportService.ts`: Vollständiger SlideMaster-Ersatz nicht möglich (API-Grenze)
- Realistischer Fallback: Farb- und Font-Mapping deck-weit via `FindReplaceService`
- `scanDeckTheme()`: Liest alle verwendeten Farben + Schriftarten
- 2 vordefinierte Presets (Office-Standard → Infront, Infront Standard → Präsentation)
- Toleranz-basierter Farb-Match

#### Schritt 11: Review / Annotationen (`src/features/review/`)
- `ReviewService.ts`: Kommentare (gelbe Textbox) + Highlights (farbiges Rechteck)
- Shape-Naming: `INFRONT_COMMENT_{timestamp}`, `INFRONT_HIGHLIGHT_{timestamp}`
- Autor-Zeitstempel-Format: `[Name – dd.mm.yyyy, hh:mm]`
- `findAllComments()`: Scannt Deck, parst Header aus Textinhalt
- `goToSlideById()`: `context.presentation.setSelectedSlides([id])` (API 1.5+)
- Einzellöschung + Bulk-Delete
- Meine-Kommentare-Ansicht mit Autor-Filter

#### Schritt 12: Gap Equalizer (`src/features/gapEqualizer/`)
- `GapEqualizerService.ts`: 3 Modi × 3 Richtungen
  - `equal`: äußere Shapes fixiert, innere gleichmäßig verteilt (min. 3)
  - `fixed`: erstes Shape fixiert, exakter Abstand (min. 2)
  - `pack`: wie `fixed` mit 0 pt (dicht packen)
- `previewGap()`: Berechnet Abstand ohne zu schreiben (Vorschau)
- Shapes per `getItemById()` geschrieben (kein redundantes load/sync)
- Undo-Snapshot vor jeder Operation

#### Schritt 13: Red Box (`src/features/redBox/`)
- `RedBoxService.ts`: INFRONT_REDBOX Safe-Area-Rahmen
- Liest Foliengröße aus `presentation.slideWidth/slideHeight` (API 1.1+)
- `ShapeFill.clear()` für transparenten Hintergrund
- `ShapeLineDashStyle` (solid/dash/dot) mit try/catch-Fallback
- Konfiguration persistent in `Document.Settings`
- `getRedBoxStatus()`: Zählt Boxen auf aktueller Folie + gesamt
- commands.ts delegiert an Service (kein duplizierter Code)

### Neu – Shared Services

- **`ConfigService`**: `getSetting<T>(key, defaultValue)`, `setSetting()`, `flushSettings()`, `loadBrandConfig()`, `saveBrandConfig()`
- **`SessionState`**: Undo-Stack (max. 10 Einträge), `createSnapshot()`, `pushUndo()`, `popUndo()`, `canUndo()`; Format-Painter-Quelle
- **`colorUtils`**: `hexToRgb()`, `rgbToHex()`, `normalizeHex()`, `colorDistance()` (euklidisch), `colorsMatch()` (Toleranz 0–30 → Schwelle 0–52), `luminance()`
- **`geometryUtils`**: `right()`, `bottom()`, `centerX()`, `centerY()`, `pxToPt()`, `ptToPx()`, `equalGap()`, `boundingBox()`, `overlapsHorizontally()`, `overlapsVertically()`
- **`logger`**: Debug/info/warn/error mit Modulpräfix; debug nur in `NODE_ENV=development`

### Shared UI-Komponenten

- **`NotificationBar`**: Wrapper um Fluent UI `MessageBar`; Types: success/warning/error/info
- **`ColorSwatch`**: Farb-Vorschau-Quadrat (22×22 pt Standard); Schachbrettmuster für transparent

### Konfiguration

- **`config/Infront_BrandConfig.json`**: Brand-Profile Default + Strict
  - Farben: Infront Navy `#003366`, Rot `#CC0000`, Weiß, Schwarz, Grau
  - Fonts: Calibri, Calibri Light
  - Toleranz: 10 (Default), 5 (Strict)

### Dokumentation

- **`README.md`**: Setup, Dev-Server, Build, Deployment, API-Limits-Übersicht
- **`TESTING.md`**: Vollständige Testdokumentation (12 Feature-Sektionen, 18 Edge-Cases, Deployment-Checkliste)
- **`SHORTCUTS.md`**: Tastaturkürzel-Dokumentation; Einschränkungen auf Mac; Accessibility-Workaround
- **`CHANGELOG.md`**: Diese Datei

### Bugfixes (innerhalb v1.0.0-Entwicklung)

- `RedBoxService.ts`: `getSetting<T>(key)` fehlte `defaultValue`-Argument → `null` übergeben
- `ReviewService.ts`: `target.setSelectedSlides()` existiert nicht auf `Slide` → `context.presentation.setSelectedSlides([target.id])`

### Bekannte Einschränkungen (Kategorie A)

| Feature | Einschränkung |
|---|---|
| Screen-Pixel-Farbpicker | EyeDropper API nicht in WKWebView/WebKit verfügbar |
| Vollständiger SlideMaster-Import | Kein `SlideMaster`-Ersatz-API in Office.js |
| Natives Undo (⌘+Z) | `Application.Undo()` nicht in Office.js; Session-State-Snapshot als Fallback |
| Globale Shortcuts | Office Add-ins dürfen keine System-Shortcuts registrieren |

### Technologie-Stack

| Paket | Version |
|---|---|
| React | 18 |
| TypeScript | 5.5 |
| Webpack | 5 |
| Fluent UI | v8 |
| office-js | CDN (Microsoft 365) |
| Office.js Requirement Set | PowerPoint 1.5+ (Basis: 1.1+) |

---

## Migrationspfad von Instrumenta (VBA)

| Instrumenta-Feature | Infront-Toolkit-Äquivalent | Status |
|---|---|---|
| Corner Radius in Pixeln | CornerRadiusPanel | ✅ implementiert |
| Screen Color Picker | ColorPickerPanel (kein Eyedropper) | ✅ Fallback implementiert |
| Brand Compliance Checker | BrandCheckPanel | ✅ implementiert |
| Format Painter Plus | FormatPainterPanel | ✅ implementiert |
| Global Find & Replace | FindReplacePanel | ✅ implementiert |
| Agenda Wizard | AgendaPanel | ✅ implementiert |
| Master Import | MasterImportPanel (Farb-/Font-Fallback) | ⚠️ API-begrenzt |
| User-Name-Stempel | ReviewPanel (Kommentare) | ✅ implementiert |
| Smart Gap Equalizer | GapEqualizerPanel | ✅ implementiert |
| Red Box | RedBoxPanel | ✅ implementiert |
| Animationen | — | ❌ außerhalb Scope |
| Slide-Sortierer | — | ❌ kein API in Office.js |
