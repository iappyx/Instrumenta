# Infront Toolkit

Professionelles PowerPoint Add-in für Strategieberater — entwickelt für **PowerPoint auf dem Mac**.

Technologie: Office.js · React · TypeScript · Swift · SwiftUI

---

## Architektur

Das Infront Toolkit besteht aus zwei Hauptkomponenten:

| Komponente | Verzeichnis | Technologie | Status |
|---|---|---|---|
| Office.js Add-in | `office-addin/` | TypeScript, React, Webpack | Fertig |
| macOS Begleit-App | `mac-app/` | Swift, SwiftUI | In Entwicklung |

```
infront-tools/
├── office-addin/      # PowerPoint Add-in (läuft im PowerPoint-WebView)
├── mac-app/           # Lokaler Server, Installationsassistent, Menu Bar App
├── scripts/           # Build- und Packaging-Skripte
├── docs/              # Dokumentation
└── legacy/            # Archivierter VBA-Code (nicht aktiv)
```

---

## Voraussetzungen

- Node.js ≥ 18, npm ≥ 9
- PowerPoint für Mac (Microsoft 365, Version ≥ 16.70 empfohlen)
- macOS 13 oder neuer
- Xcode 15+ (nur für mac-app Build)

---

## Lokales Setup (Entwicklung)

```bash
# Abhängigkeiten installieren (Add-in)
sh scripts/setup-dev.sh

# oder manuell:
cd office-addin && npm install
```

### Dev-Server starten

```bash
cd office-addin
npm run dev-server
# → https://localhost:3000
```

**HTTPS-Zertifikat vertrauen (einmalig):**
```bash
open "https://localhost:3000"
# Safari/Chrome-Dialog: "Vertrauen" wählen
```

### Add-in in PowerPoint sideloaden (Entwicklung)

1. PowerPoint öffnen
2. **Einfügen → Add-ins → Meine Add-ins → Add-in hochladen**
3. `office-addin/manifest.xml` auswählen
4. Tab **Infront Toolkit** erscheint im Ribbon

---

## Build

```bash
# Production Build (office-addin/dist/)
sh scripts/build-addin.sh

# oder manuell:
cd office-addin && npm run build
```

### Typecheck & Lint

```bash
cd office-addin
npm run typecheck   # TypeScript prüfen
npm run lint        # ESLint
```

---

## Projektstruktur (Add-in)

```
office-addin/
├── manifest.xml              # Office Add-in Manifest (GUID, Ribbon, URLs)
├── package.json
├── tsconfig.json
├── webpack.config.js
├── assets/icons/             # Add-in Icons (16/32/80 px)
├── config/
│   └── Infront_BrandConfig.json   # Brand-Profile (Default + Strict)
└── src/
    ├── commands/             # ExecuteFunction Ribbon-Handler
    ├── taskpane/             # React Task Pane App
    │   ├── App.tsx           # Root + URL-Router (?view=)
    │   └── components/       # Feature-Panel-Komponenten
    ├── features/             # Feature-Services (10 Module)
    ├── services/             # Office.js Abstraktionsschicht + Config + State
    └── utils/                # Farb-, Geometrie-, Logger-Utilities
```

---

## Ribbon-Struktur

| Gruppe | Button | Typ |
|---|---|---|
| Shapes | Eckenradius | Task Pane |
| Format | Farbwähler, Format Painter+ | Task Pane |
| Quality | Brand Check, Suchen & Ersetzen, Kommentare entfernen, Red Box entfernen | Task Pane / Direkt |
| Ausrichten | Gap H, Gap V, Gap... | Direkt / Task Pane |
| Struktur | Agenda | Task Pane |
| Design | Master importieren, Red Box, Red Box: Alle Slides, Red Box Einstellungen | Task Pane / Direkt |
| Review | Kommentar, Markieren, Meine Kommentare | Task Pane / Direkt |

---

## Bekannte Einschränkungen (Mac / Office.js)

Vollständige Liste: [TESTING.md](TESTING.md) · Architekturentscheidungen: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

| Einschränkung | Grund | Fallback |
|---|---|---|
| Screen-Pixel-Picker | WebKit/Safari unterstützt EyeDropper API nicht | Hex-Eingabe + Shape-Farbübernahme |
| Master vollständig ersetzen | Kein SlideMaster-Replacement in Office.js | Theme-Farben/-Fonts anpassen |
| Natives Undo | Kein `Application.Undo()` in Office.js | Session-Snapshot + Revert-Button |
| Globale Shortcuts | Office Add-ins können keine globalen Shortcuts registrieren | Ribbon-Buttons als Primärpfad |

---

## Installation (Endnutzer)

Vollständige Installationsanleitung: [docs/INSTALL_MAC.md](docs/INSTALL_MAC.md)

Kurzübersicht:
1. `InfrontToolkit.dmg` herunterladen und öffnen
2. `InfrontToolkit.app` in `/Applications` ziehen
3. App starten → Setup-Assistent führt durch die Einrichtung
4. PowerPoint öffnen → Infront-Tab im Ribbon

---

## Release & Packaging

Vollständiger Release-Prozess: [docs/RELEASE.md](docs/RELEASE.md)

```bash
sh scripts/build-addin.sh     # Add-in bauen
sh scripts/build-mac-app.sh   # macOS App bauen (Schritt 3)
sh scripts/package-dmg.sh     # DMG erstellen (Schritt 7)
```

---

## Legacy-Referenz

`legacy/` enthält den ursprünglichen VBA-Code (Instrumenta) sowie alte Installer.
Diese Dateien sind **nicht aktiv** und dienen nur als Feature-Referenz.
Siehe: `legacy/vba-source/` für den VBA-Code.
