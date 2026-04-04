#!/usr/bin/env bash
# build-mac-app.sh
# Baut die native macOS Begleit-App (InfrontToolkit.app).
#
# Voraussetzungen:
#   - Xcode 15+ (macOS)
#   - xcodebuild im PATH (enthalten in Xcode Command Line Tools)
#   - Add-in muss vorher gebaut worden sein: sh scripts/build-addin.sh
#
# Code Signing (Produktion):
#   - Apple Developer Account + Developer ID Application Certificate
#   - Für Entwicklung: CODE_SIGNING_ALLOWED=NO überspringt Signierung
#
# Optionale Abhängigkeit:
#   - xcodegen (brew install xcodegen): Regeneriert project.pbxproj aus project.yml

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
MAC_APP_DIR="$REPO_ROOT/mac-app"
BUILD_DIR="$REPO_ROOT/.build/mac-app"
XCODEPROJ="$MAC_APP_DIR/InfrontToolkit.xcodeproj"

echo "==> [build-mac-app] Native macOS App Build"
echo "    Repository: $REPO_ROOT"
echo "    Projekt:    $XCODEPROJ"
echo ""

# -------------------------------------------------------
# 0. Voraussetzungen prüfen
# -------------------------------------------------------
if ! command -v xcodebuild &>/dev/null; then
  echo "    FEHLER: xcodebuild nicht gefunden."
  echo "    Installiere Xcode Command Line Tools: xcode-select --install"
  exit 1
fi

XCODE_VERSION=$(xcodebuild -version 2>/dev/null | head -1 || echo "unbekannt")
echo "==> [0/4] Xcode: $XCODE_VERSION"

# -------------------------------------------------------
# 1. Add-in dist/ kopieren (falls noch nicht geschehen)
# -------------------------------------------------------
ADDIN_DIST="$REPO_ROOT/office-addin/dist"
APP_RESOURCES="$MAC_APP_DIR/InfrontToolkit/Resources/addin-dist"

echo "==> [1/4] Add-in Assets einbinden..."
if [ ! -d "$ADDIN_DIST" ]; then
  echo "    FEHLER: office-addin/dist/ fehlt."
  echo "    Bitte zuerst ausführen: sh scripts/build-addin.sh"
  exit 1
fi

rm -rf "$APP_RESOURCES"
cp -r "$ADDIN_DIST" "$APP_RESOURCES"
echo "    OK: $ADDIN_DIST → $APP_RESOURCES"

# -------------------------------------------------------
# 2. Optional: XcodeGen (regeneriert project.pbxproj)
# -------------------------------------------------------
echo "==> [2/4] Xcode Projekt vorbereiten..."
if command -v xcodegen &>/dev/null; then
  echo "    xcodegen gefunden – regeneriere project.pbxproj aus project.yml..."
  cd "$MAC_APP_DIR" && xcodegen generate --quiet
  echo "    OK"
else
  echo "    xcodegen nicht installiert – verwende bestehendes project.pbxproj"
  echo "    (Tipp: brew install xcodegen für sauberere Projektverwaltung)"
fi

# -------------------------------------------------------
# 3. App bauen mit xcodebuild
# -------------------------------------------------------
echo "==> [3/4] xcodebuild Release..."
mkdir -p "$BUILD_DIR"

xcodebuild \
  -project "$XCODEPROJ" \
  -scheme InfrontToolkit \
  -configuration Release \
  -derivedDataPath "$BUILD_DIR" \
  CODE_SIGNING_ALLOWED=NO \
  clean build 2>&1 | grep -E "(error:|warning:|BUILD (SUCCEEDED|FAILED)|^    /)" || true

# Build-Ergebnis prüfen
APP_PATH="$BUILD_DIR/Build/Products/Release/InfrontToolkit.app"
if [ ! -d "$APP_PATH" ]; then
  # Alternativer Pfad
  APP_PATH=$(find "$BUILD_DIR" -name "InfrontToolkit.app" -type d 2>/dev/null | head -1 || true)
fi

if [ -z "$APP_PATH" ] || [ ! -d "$APP_PATH" ]; then
  echo ""
  echo "    FEHLER: InfrontToolkit.app nicht gefunden nach Build."
  echo "    Vollständige xcodebuild Ausgabe: $BUILD_DIR"
  exit 1
fi

echo "    OK: $APP_PATH"

# -------------------------------------------------------
# 4. Zusammenfassung
# -------------------------------------------------------
echo ""
echo "==> [4/4] Build abgeschlossen."
echo "    App:     $APP_PATH"
echo "    Größe:   $(du -sh "$APP_PATH" | cut -f1)"
echo ""
echo "Nächste Schritte:"
echo "  • Zum Testen: open \"$APP_PATH\""
echo "  • DMG erstellen: sh scripts/package-dmg.sh"
echo ""
echo "HINWEIS: CODE_SIGNING_ALLOWED=NO – die App ist NICHT signiert."
echo "         Für Produktion: Apple Developer ID Certificate + Notarization benötigt."
echo "         Siehe: docs/RELEASE.md"
