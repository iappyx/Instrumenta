# Infront Toolkit

Professionelles PowerPoint Add-in für Strategieberater — entwickelt für **PowerPoint auf dem Mac**.

Technologie: Office.js · React · TypeScript · Swift · SwiftUI

---

## Architektur

Das Infront Toolkit besteht aus zwei Hauptkomponenten:

| Komponente | Verzeichnis | Technologie |
|---|---|---|
| Office.js Add-in | `office-addin/` | TypeScript, React, Webpack |
| macOS Begleit-App | `mac-app/` | Swift, SwiftUI, Network.framework |

```
infront-tools/
├── office-addin/      # PowerPoint Add-in (läuft im PowerPoint-WebView)
├── mac-app/           # Lokaler Server, Setup-Assistent, Menu Bar App
├── scripts/           # Build- und Packaging-Skripte
├── docs/              # Dokumentation
├── VERSION            # Versionsnummer (Single Source of Truth)
└── legacy/            # Archivierter VBA-Code (nicht aktiv)
```

**Kommunikation:** Das Add-in lädt seine Assets und Konfiguration vom lokalen HTTP-Server der macOS App (`http://localhost:3000`). Keine Cloud-Abhängigkeit.

---

## Voraussetzungen

- Node.js ≥ 18, npm ≥ 9
- PowerPoint für Mac (Microsoft 365)
- macOS 13 Ventura oder neuer
- Xcode 15+ (nur für mac-app Build)

---

## Installation (Endnutzer)

Vollständige Anleitung: [docs/INSTALL_MAC.md](docs/INSTALL_MAC.md)

```
1. InfrontToolkit.dmg öffnen → App in /Applications ziehen
2. App starten → Setup-Assistent (ca. 5 Min.)
3. PowerPoint öffnen → Tab „Infront Toolkit" im Ribbon
```

---

## Lokale Entwicklung

```bash
# Abhängigkeiten installieren
sh scripts/setup-dev.sh

# Dev-Server starten (Add-in mit Hot Reload)
cd office-addin && npm run dev-server
# → https://localhost:3000

# Add-in in PowerPoint sideloaden (Entwicklung)
# Einfügen → Add-ins → Add-in hochladen → office-addin/manifest.xml
```

**HTTPS-Zertifikat vertrauen (einmalig):**
```bash
open "https://localhost:3000"
# Im Browser: „Vertrauen" / Zertifikat akzeptieren
```

---

## Build

### Add-in

```bash
sh scripts/build-addin.sh
# → office-addin/dist/

# Typecheck + Lint
cd office-addin
npm run typecheck
npm run lint
```

### macOS App

```bash
sh scripts/build-mac-app.sh
# → .build/mac-app/Build/Products/Release/InfrontToolkit.app
```

### Vollst��ndiger Release-Build (Add-in + App + DMG)

```bash
# Entwicklung (kein Zertifikat nötig)
SKIP_SIGNING=1 sh scripts/build-release.sh

# Produktion (mit Code Signing + Notarization)
CODESIGN_IDENTITY="Developer ID Application: ..." \
NOTARIZATION_PROFILE="InfrontToolkitNotarization" \
sh scripts/build-release.sh
# → .build/dist/InfrontToolkit-<version>.dmg
```

Release-Prozess: [docs/RELEASE.md](docs/RELEASE.md)

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
    │   ├── index.tsx         # Einstiegspunkt + macOS-Integration
    │   ├── App.tsx           # Root + URL-Router (?view=)
    │   └── components/       # Feature-Panel-Komponenten
    ├── features/             # Feature-Services (10 Module)
    ├── services/             # Office.js Abstraktionsschicht + Config + State
    └── utils/                # Farb-, Geometrie-, Logger-Utilities
```

## Projektstruktur (macOS App)

```
mac-app/InfrontToolkit/
├── InfrontToolkitApp.swift       # App Entry Point (MenuBarExtra, Window)
├── Models/
│   ├── AppState.swift            # @MainActor ObservableObject (Server, Logs, Setup)
│   └── ServerConfig.swift        # Port, addinDistURL, Verfügbarkeitsprüfung
├── Services/
│   ├── LocalServerService.swift  # HTTP-Server (NWListener, 127.0.0.1)
│   ├── ConfigSyncService.swift   # Brand Config (User-Override → Bundle → Fallback)
│   ├── ManifestService.swift     # Manifest aus Template generieren + installieren
│   └── AppAutoStartService.swift # Login Item (SMAppService)
└── Views/
    ├── StatusMenuView.swift       # Menu Bar Dropdown
    ├── SetupWizardView.swift      # 6-Schritte Erststart-Assistent
    ├── MainWindowView.swift       # Dashboard (Status / Einstellungen / Logs)
    └── SettingsView.swift         # Port, Autostart, Manifest, Brand Config
```

---

## HTTP-Endpunkte (LocalServerService)

| Endpoint | Methode | Beschreibung |
|---|---|---|
| `/health` | GET | `{"status":"ok","version":"1.1.0","addinDistPresent":bool}` |
| `/config/Infront_BrandConfig.json` | GET | Brand Config (User-Override oder Bundle-Default) |
| `/log` | POST | Log-Eintrag vom Add-in → Logs-Tab in der App |
| `/*` | GET | Statische Dateien aus `addin-dist/` (Add-in Assets) |

Details: [docs/INTEGRATION.md](docs/INTEGRATION.md)

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

Vollständige Liste: [TESTING.md](TESTING.md) · Architektur: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

| Einschränkung | Grund | Fallback |
|---|---|---|
| Screen-Pixel-Picker | WebKit unterstützt EyeDropper API nicht | Hex-Eingabe + Shape-Farbübernahme |
| Master vollständig ersetzen | Kein SlideMaster-Replacement in Office.js | Theme-Farben/-Fonts anpassen |
| Natives Undo | Kein `Application.Undo()` in Office.js | Session-Snapshot + Revert-Button |
| Globale Shortcuts | Office Add-ins können keine Systemshortcuts registrieren | Ribbon-Buttons als Primärpfad |

---

## Dokumentation

| Dokument | Inhalt |
|---|---|
| [docs/INSTALL_MAC.md](docs/INSTALL_MAC.md) | Endnutzer-Installationsanleitung |
| [docs/RELEASE.md](docs/RELEASE.md) | Release-Prozess, Code Signing, Notarization |
| [docs/INTEGRATION.md](docs/INTEGRATION.md) | HTTP-Endpunkte, Kommunikationsmodell |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Architekturentscheidungen, Komponentenverantwortung |
| [TESTING.md](TESTING.md) | Testchecklisten, API-Limits, Edge Cases |
| [CHANGELOG.md](CHANGELOG.md) | Versionshistorie |

---

## Legacy-Referenz

`legacy/` enthält den ursprünglichen VBA-Code (Instrumenta) sowie alte Installer.
Diese Dateien sind **nicht aktiv** und dienen nur als Feature-Referenz.
