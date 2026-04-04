# InfrontToolkit – macOS Begleit-App

Native macOS Menu Bar App für das Infront Toolkit PowerPoint Add-in.

## Aufgaben der App

| Funktion | Beschreibung |
|---|---|
| Lokaler HTTP-Server | Serviert `addin-dist/` auf `http://localhost:3000` |
| Setup-Assistent | Führt durch Ersteinrichtung (Port → Server → Manifest → Office → Autostart) |
| Manifest-Installation | Schreibt `manifest.xml` in den Office Shared Folder Catalog |
| Autostart | Registriert sich als macOS Login Item (SMAppService) |
| Status-Anzeige | Menu Bar Icon zeigt Serverstatus (läuft / gestoppt / Fehler) |

## Voraussetzungen

- macOS 13.0 (Ventura) oder neuer
- Xcode 15+ mit macOS SDK
- Für Code Signing / Notarization: Apple Developer Account

## Projekt öffnen

```bash
# Option A: Xcode Projekt direkt öffnen (project.pbxproj)
open mac-app/InfrontToolkit.xcodeproj

# Option B: Mit XcodeGen neu generieren (empfohlen nach Änderungen)
brew install xcodegen   # einmalig
cd mac-app && xcodegen generate
open InfrontToolkit.xcodeproj
```

## Build

```bash
# Debug Build (Entwicklung)
xcodebuild -project InfrontToolkit.xcodeproj \
           -scheme InfrontToolkit \
           -configuration Debug \
           build

# Release Build
xcodebuild -project InfrontToolkit.xcodeproj \
           -scheme InfrontToolkit \
           -configuration Release \
           -derivedDataPath ../.build/mac-app \
           clean build

# Oder über scripts/:
cd .. && sh scripts/build-mac-app.sh
```

## Add-in Assets einbinden

Die App benötigt den gebauten Add-in `dist/` unter:
```
mac-app/InfrontToolkit/Resources/addin-dist/
```

Dieses Verzeichnis wird automatisch befüllt durch:
```bash
cd .. && sh scripts/build-addin.sh   # Baut office-addin/dist/
         sh scripts/copy-dist-to-app.sh  # Kopiert dist/ → Resources/addin-dist/
```

## Projektstruktur

```
mac-app/
├── InfrontToolkit.xcodeproj/      # Xcode Projektdatei
├── project.yml                    # XcodeGen Konfiguration (Alternative)
└── InfrontToolkit/
    ├── InfrontToolkitApp.swift    # @main Einstiegspunkt (MenuBarExtra)
    ├── Models/
    │   ├── AppState.swift         # ObservableObject: Status, Config, Logs
    │   └── ServerConfig.swift     # Port, Pfade, addin-dist Prüfung
    ├── Services/
    │   ├── LocalServerService.swift    # NWListener HTTP-Server
    │   ├── ManifestService.swift       # Manifest installieren/entfernen
    │   └── AppAutoStartService.swift   # SMAppService Login Item
    ├── Views/
    │   ├── StatusMenuView.swift        # Menu Bar Dropdown
    │   ├── SetupWizardView.swift       # 6-Schritte Erststart-Assistent
    │   ├── MainWindowView.swift        # Dashboard (Status + Logs + Settings)
    │   └── SettingsView.swift          # Port, Autostart, Manifest
    ├── Resources/
    │   ├── manifest-template.xml       # Manifest-Vorlage ({{PORT}} Platzhalter)
    │   └── addin-dist/                 # Add-in Assets (von build-addin.sh befüllt)
    ├── Assets.xcassets/               # App Icon (noch Platzhalter)
    ├── Info.plist                     # LSUIElement=YES (kein Dock-Icon)
    └── InfrontToolkit.entitlements    # App Sandbox, Netzwerk, Dateisystem
```

## Bekannte Einschränkungen

| Einschränkung | Details |
|---|---|
| App Sandbox | Netzwerk-Server in Sandbox erfordert `com.apple.security.network.server` Entitlement ✓ gesetzt |
| Autostart | `SMAppService` erfordert korrekte Code-Signatur; schlägt im Development-Build ohne Zertifikat fehl |
| Code Signing | Apple Developer ID Certificate erforderlich für Produktion (Gatekeeper) |
| Notarization | `xcrun notarytool` + Apple Developer Account für öffentliche Distribution |
| addin-dist leer | App zeigt Fehlermeldung wenn addin-dist/ fehlt oder leer ist; `build-addin.sh` ausführen |
