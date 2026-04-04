# Infront Toolkit – Release-Prozess

> Voraussetzungen: macOS 13+, Xcode 15+, Node.js 18+, Apple Developer Account (für Produktion)

---

## Übersicht

```
VERSION Datei     ←  Versionsnummer (Single Source of Truth)
      │
      ▼
build-addin.sh    →  office-addin/dist/
      │
      ▼
copy-dist-to-app.sh  →  mac-app/Resources/addin-dist/
      │
      ▼
build-mac-app.sh  →  .build/mac-app/Build/Products/Release/InfrontToolkit.app
      │
      ▼
package-dmg.sh    →  .build/dist/InfrontToolkit-<version>.dmg
```

---

## Schritt 1 – Version aktualisieren

**`VERSION`** (Repo-Root) ist die einzige Datei die angepasst wird:

```
1.1.0
```

Dieses Feld wird von `package-dmg.sh` und `build-release.sh` gelesen.

Außerdem manuell aktualisieren:
- `office-addin/package.json` → `"version": "..."`
- `mac-app/InfrontToolkit/Services/LocalServerService.swift` → JSON in `buildHealthResponse()` (`"version":"..."`)

---

## Schritt 2 – Voraussetzungen prüfen

```bash
# Xcode Command Line Tools
xcodebuild -version
# → Xcode 15.x o. ä.

# Node.js
node --version
# → v18.x oder neuer

# Optional: Hübscheres DMG
brew install create-dmg

# Code Signing (Produktion)
security find-identity -v -p codesigning | grep "Developer ID Application"
# → 1 valid identity
```

---

## Schritt 3 – Code Signing einrichten (einmalig, Produktion)

### 3a. Apple Developer Account

1. [developer.apple.com](https://developer.apple.com) → Certificates, IDs & Profiles
2. Neues Zertifikat: **Developer ID Application**
3. Herunterladen + ins Keychain importieren

### 3b. Notarization Credentials speichern (einmalig)

```bash
xcrun notarytool store-credentials "InfrontToolkitNotarization" \
  --apple-id "dein@apple.id" \
  --team-id "DEIN_TEAM_ID" \
  --password "app-specific-password"
```

App-spezifisches Passwort erstellen: [appleid.apple.com](https://appleid.apple.com) → App-spezifische Passwörter

### 3c. Identity herausfinden

```bash
security find-identity -v -p codesigning
# Beispiel-Output:
# 1) A1B2C3D4E5F6... "Developer ID Application: Infront AG (ABCDEF1234)"
```

---

## Schritt 4 – Release Build

### Development (ohne Signierung)

```bash
SKIP_SIGNING=1 sh scripts/build-release.sh
```

### Produktion (mit Signierung + Notarization)

```bash
export CODESIGN_IDENTITY="Developer ID Application: Infront AG (ABCDEF1234)"
export NOTARIZATION_PROFILE="InfrontToolkitNotarization"

sh scripts/build-release.sh
```

### Einzelne Schritte separat

```bash
# Nur Add-in
sh scripts/build-addin.sh

# Nur macOS App (mit Signierung)
sh scripts/build-mac-app.sh

# Nur DMG (App muss vorher gebaut worden sein)
CODESIGN_IDENTITY="..." NOTARIZATION_PROFILE="..." sh scripts/package-dmg.sh
```

---

## Schritt 5 – DMG testen

```bash
VERSION=$(cat VERSION | tr -d '[:space:]')
open ".build/dist/InfrontToolkit-${VERSION}.dmg"
```

**Checkliste:**
- [ ] DMG öffnet sich korrekt
- [ ] App-Icon + /Applications-Symlink sichtbar
- [ ] App in /Applications kopieren und starten
- [ ] Menu Bar Icon erscheint (⚡)
- [ ] Setup-Assistent startet bei Erststart
- [ ] Server kann gestartet werden
- [ ] Manifest wird installiert
- [ ] PowerPoint-Add-in lädt nach Office-Catalog-Setup

---

## Schritt 6 – Notarization prüfen (Produktion)

```bash
VERSION=$(cat VERSION | tr -d '[:space:]')
spctl -a -vvv -t open ".build/dist/InfrontToolkit-${VERSION}.dmg"
# → accepted  source=Notarized Developer ID

# App-Signierung prüfen
codesign --verify --deep --strict --verbose /Applications/InfrontToolkit.app
# → InfrontToolkit.app: valid on disk
```

---

## Schritt 7 – GitHub Release anlegen

```bash
VERSION=$(cat VERSION | tr -d '[:space:]')

# Tag anlegen
git tag -a "v${VERSION}" -m "Release v${VERSION}"
git push origin "v${VERSION}"

# GitHub Release (manuell über github.com/joranpauleit/infront-tools/releases/new)
# oder via gh CLI:
gh release create "v${VERSION}" \
  ".build/dist/InfrontToolkit-${VERSION}.dmg" \
  --title "Infront Toolkit v${VERSION}" \
  --notes "$(cat CHANGELOG.md | awk '/^## \['"${VERSION}"'\]/{found=1; next} /^## \[/{if(found) exit} found{print}')"
```

---

## Schritt 8 – Post-Release Checkliste

- [ ] Git-Tag auf Remote gepusht
- [ ] GitHub Release mit DMG-Anhang erstellt
- [ ] CHANGELOG.md aktualisiert
- [ ] VERSION-Datei für nächsten Entwicklungszyklus angehoben (z.B. `1.2.0-dev`)
- [ ] INSTALL_MAC.md Versionsnummer geprüft (ist versionsneutral formuliert – OK)

---

## Fehler beim Build

### xcodebuild schlägt fehl

```bash
# Vollständige Ausgabe ansehen
xcodebuild \
  -project mac-app/InfrontToolkit.xcodeproj \
  -scheme InfrontToolkit \
  -configuration Release \
  CODE_SIGNING_ALLOWED=NO \
  clean build 2>&1 | tee /tmp/xcodebuild.log
```

### Notarization schlägt fehl

```bash
# Status eines laufenden Uploads prüfen
xcrun notarytool log <submission-id> --keychain-profile InfrontToolkitNotarization
```

Häufige Ursachen:
- **Hardened Runtime** nicht aktiviert → `ENABLE_HARDENED_RUNTIME=YES` in Build Settings
- **Entitlements** fehlen → `InfrontToolkit.entitlements` prüfen
- **Netzwerkzugriff** nicht erlaubt → `com.apple.security.network.server` in Entitlements

### Gatekeeper-Warnung beim Testen (Development)

```bash
# Quarantäne-Flag entfernen (lokal)
xattr -cr /Applications/InfrontToolkit.app
```

---

## Versionierungsschema

```
MAJOR.MINOR.PATCH[-label]

1.0.0         → Erste stabile Version
1.1.0-alpha   → Neues Feature (unstable)
1.1.0-beta    → Feature complete, testing
1.1.0-rc1     → Release Candidate
1.1.0         → Stable Release
```
