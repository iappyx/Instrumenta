# Infront Toolkit – Architektur-Dokumentation

> Erstellt: Schritt 2 der macOS-Migration (2026-04-04)

---

## Übersicht

Das Infront Toolkit besteht nach der Migration aus drei Hauptkomponenten:

```
infront-tools/
├── office-addin/     # Office.js Add-in (TypeScript/React) – läuft in PowerPoint
├── mac-app/          # Native macOS Begleit-App (Swift/SwiftUI) – [Schritt 3]
├── scripts/          # Build- und Packaging-Skripte
├── docs/             # Dokumentation
└── legacy/           # Archivierter VBA-Code (nicht aktiv)
```

---

## Komponentenverantwortung

### `office-addin/` — Office.js Add-in

**Was hier bleibt (muss Office.js bleiben):**

Alle Funktionen, die direkt mit PowerPoint-Objekten arbeiten, können **ausschließlich** über die Office JavaScript API realisiert werden. Sie laufen im WKWebView-Browser von PowerPoint und haben keinen direkten Zugriff auf macOS-APIs.

| Modul | Verantwortung |
|-------|---------------|
| `src/features/*/` | Feature-Services: Shapes lesen/schreiben, Slides manipulieren |
| `src/services/powerpoint/` | Office.js Abstraktionsschicht (Shape, Slide, Selection, Text) |
| `src/taskpane/` | React-UI im PowerPoint Task Pane |
| `src/commands/commands.ts` | Ribbon ExecuteFunction-Handler |
| `manifest.xml` | Office Add-in Manifest (Pflicht) |

**Technologie-Stack:**
- React 18 + TypeScript 5.5 (strict)
- Webpack 5 (Build), webpack-dev-server (Development)
- Fluent UI v8 (UI-Komponenten)
- office-js vom CDN (`appsforoffice.microsoft.com`)

**Build:**
```bash
cd office-addin
npm install
npm run build        # → office-addin/dist/
npm run dev-server   # → https://localhost:3000 (Entwicklung)
```

---

### `mac-app/` — Native macOS Begleit-App [Schritt 3]

**Was hier sinnvoll ist (kann/soll nativ sein):**

| Funktion | Begründung |
|----------|-----------|
| Lokaler HTTP-Server (port 3000) | Dient `dist/` Assets lokal aus; ersetzt den webpack-dev-server für Produktion |
| Manifest-Catalog-Setup | Konfiguriert Office Shared Folder Catalog – macOS-Dateisystem-Operation |
| Setup-Assistent | Führt Nutzer durch Ersteinrichtung – kein PowerPoint-API nötig |
| Autostart (Login Item) | `SMAppService` – macOS-native API |
| Status-Anzeige (Menu Bar) | Serverstatus, Port, Fehler – nativ in AppKit/SwiftUI |
| Brand-Config-Editor | Lesen/Schreiben von `Infront_BrandConfig.json` – Dateisystemzugriff |
| Logging | Lokale Log-Datei, Anzeige in App |

**Was NICHT in die native App gehört:**
- Jegliche PowerPoint-Objektmanipulation → muss in `office-addin/` bleiben
- Office.js API-Aufrufe → nicht außerhalb des WebView möglich
- Ribbon-Definitionen → gehören in `manifest.xml`

**Technologie:**
- Swift 5.9+
- SwiftUI für UI-Komponenten
- AppKit für Menu Bar (NSStatusItem)
- Network.framework für lokalen HTTP-Server
- Foundation für Dateisystem, JSON, Logging

---

### Kommunikation zwischen Add-in und nativer App

Die Kommunikation ist bewusst **minimal und lose gekoppelt**:

```
PowerPoint (WebView)          macOS App
      │                            │
      │   HTTP GET /config.json    │
      │──────────────────────────>│
      │   ← Brand Config JSON      │
      │                            │
      │   HTTP POST /log           │  (optional, Schritt 5)
      │──────────────────────────>│
```

- Das Add-in lädt `Infront_BrandConfig.json` über HTTP vom lokalen Server
- Der lokale Server antwortet auf `http://localhost:3000/`
- Keine bidirektionale Socket-Verbindung, keine nativen Bridge-APIs
- Keine `window.webkit.messageHandlers` (fragil, Office-interne WebView nicht kontrollierbar)

**Sicherheitsabgrenzung:**
- Der lokale HTTP-Server lauscht nur auf `127.0.0.1` (loopback), nicht auf externe Interfaces
- Kein CORS-Zugriff von externen Origins

---

## Installations-Architektur

### Deployment-Modell: Lokaler HTTP-Server (Offline-fähig)

Office für Mac erlaubt `http://localhost` als SourceLocation (Ausnahme für lokale Entwicklung und Verteilung). Dies ist der robusteste Weg für eine DMG-Distribution ohne Cloud-Hosting.

```
DMG
 └── InfrontToolkit.app
      ├── Contents/MacOS/InfrontToolkit   (ausführbare Datei)
      └── Contents/Resources/
           ├── addin-dist/                (gebundelte dist/ vom Build)
           │   ├── taskpane.html
           │   ├── taskpane.js
           │   ├── commands.html
           │   ├── commands.js
           │   └── assets/icons/
           ├── manifest-template.xml      (Vorlage mit {{PORT}} Platzhalter)
           └── Infront_BrandConfig.json   (Standard-Konfiguration)
```

### Installationsfluss (Endnutzer)

```
1. DMG öffnen → InfrontToolkit.app in /Applications ziehen
2. App starten → Setup-Assistent erscheint
3. Setup-Assistent:
   a. Port wählen (Default: 3000)
   b. Manifest in ~/Library/Group Containers/.../Manifests/ ablegen
      ODER Anleitung zum manuellen Catalog-Setup zeigen
   c. Autostart als Login Item aktivieren (optional)
   d. Test-Button: "Add-in erreichbar?" (HTTP GET auf localhost:3000)
4. App läuft als Menu Bar App (dauerhaft im Hintergrund)
5. Nutzer öffnet PowerPoint:
   - Add-in erscheint im Ribbon (wenn Manifest korrekt geladen)
   - Task Pane lädt von http://localhost:3000/taskpane.html
```

### Office Manifest-Registrierung (Realitäts-Check)

**Option A: Shared Folder Catalog (empfohlen)**
- Nutzer konfiguriert einmalig einen Ordnerpfad in Office-Einstellungen
- Word/Excel/PowerPoint → Menü → Add-ins → Weitere Add-ins → Freigegebene Ordnerkataloge
- Pfad: der Ordner, in dem die macOS App manifest.xml ablegt
- Vorteil: keine Admin-Rechte nötig, funktioniert lokal
- Nachteil: einmalige manuelle Aktion des Nutzers

**Option B: Sideload (Entwicklung)**
- Einfügen → Add-ins → Meine Add-ins → Add-in hochladen → manifest.xml auswählen
- Nur für Entwicklung und Tests geeignet

**Option C: Microsoft 365 Admin Center (Enterprise)**
- IT-Administrator deployed Manifest für alle Nutzer
- Für Unternehmensverteilung am saubersten
- Kein DMG nötig, nur Manifest-URL

---

## Build-Chain Übersicht

```
scripts/setup-dev.sh          → Entwicklungsumgebung einrichten
scripts/build-addin.sh        → office-addin/dist/ erstellen
scripts/build-mac-app.sh      → mac-app/InfrontToolkit.app erstellen [Schritt 3]
scripts/copy-dist-to-app.sh   → dist/ → mac-app/Resources/addin-dist/ kopieren
scripts/package-dmg.sh        → InfrontToolkit.dmg erstellen [Schritt 7]
```

Vollständiger Release-Build (Schritt 7):
```bash
sh scripts/build-addin.sh
sh scripts/build-mac-app.sh
sh scripts/package-dmg.sh
```

---

## Code Signing & Notarization (Produktions-Anforderungen)

> **Explizite Einschränkung**: Diese Punkte sind in der aktuellen Entwicklungsphase
> NICHT vollständig implementiert. Ein Apple Developer Account ist Voraussetzung.

| Anforderung | Beschreibung | Status |
|-------------|--------------|--------|
| Apple Developer Account | $99/Jahr, benötigt für Zertifikate | ❌ Nicht eingerichtet |
| Developer ID Application Cert | Für Code Signing der .app | ❌ Nicht vorhanden |
| Notarization | `xcrun notarytool` – verhindert Gatekeeper-Block | ❌ Nicht implementiert |
| Hardened Runtime | Voraussetzung für Notarization | ⏳ Schritt 7 |

Ohne Code Signing + Notarization erscheint beim ersten Start die macOS Gatekeeper-Warnung.
Nutzer können die App manuell freigeben über: Systemeinstellungen → Datenschutz & Sicherheit.

---

## Abhängigkeitsübersicht

### Office.js Add-in
| Abhängigkeit | Version | Zweck |
|---|---|---|
| react | 18.3.1 | UI-Framework |
| react-dom | 18.3.1 | DOM-Rendering |
| @fluentui/react | 8.120.3 | Microsoft UI-Komponenten |
| typescript | 5.5.4 | Type Safety |
| webpack | 5.94.0 | Build |
| webpack-dev-server | 5.0.4 | Dev-Server |

### Native macOS App (geplant, Schritt 3)
- Swift 5.9+ (Standardbibliothek)
- SwiftUI (macOS 13+)
- AppKit (NSStatusItem für Menu Bar)
- Network.framework (lokaler HTTP-Server)
- Keine Drittanbieter-Abhängigkeiten (Ziel)

---

## Bekannte architektonische Grenzen

1. **Office.js ist the single source of truth für PowerPoint-Objekte** — kein nativer Zugriff möglich
2. **Local HTTP statt HTTPS** — für localhost in Office akzeptiert, nicht für externe Deployments
3. **Manifest Sideloading ≠ echter App Store** — Shared Folder Catalog erfordert einmalige manuelle Office-Konfiguration
4. **App muss laufen** — solange PowerPoint das Add-in nutzt, muss der lokale Server aktiv sein
5. **No EyeDropper in WKWebView** — Screen-Pixel-Picker nicht realisierbar (WebKit-Einschränkung)
6. **SlideMaster-Ersatz nicht möglich** — Office.js API-Grenze, dokumentiert in TESTING.md
