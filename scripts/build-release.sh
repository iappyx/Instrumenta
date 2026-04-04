#!/usr/bin/env bash
# build-release.sh
# Vollständiger Release-Build: Add-in → macOS App → DMG
#
# Führt alle Build-Schritte in der korrekten Reihenfolge aus:
#   1. Add-in bauen (npm run build)
#   2. dist/ in App Resources kopieren
#   3. macOS App bauen (xcodebuild)
#   4. DMG erstellen
#
# Umgebungsvariablen (optional):
#   CODESIGN_IDENTITY    – Code-Signing: "Developer ID Application: ..."
#   NOTARIZATION_PROFILE – notarytool keychain profile
#   SKIP_SIGNING=1       – Signierung überspringen (Development/CI)
#
# Verwendung:
#   sh scripts/build-release.sh
#   SKIP_SIGNING=1 sh scripts/build-release.sh
#   CODESIGN_IDENTITY="Developer ID Application: ..." sh scripts/build-release.sh

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="$(cat "$REPO_ROOT/VERSION" | tr -d '[:space:]')"
SCRIPTS="$REPO_ROOT/scripts"

echo "╔════════════════════���═════════════════════���═══════════════╗"
echo "║  Infront Toolkit – Release Build v${VERSION}            "
echo "╚══════════════════════════════════════��═══════════════════╝"
echo ""

START_TIME=$(date +%s)

# -------------------------------------------------------
# Schritt 1: Add-in bauen
# -------------------------------------------------------
echo "━━━ [1/4] Add-in Build (npm) ━━━━━━━━━━━━━━━━���━━━━━━━━━━━━━"
sh "$SCRIPTS/build-addin.sh"
echo ""

# -------------------------------------------------------
# Schritt 2: dist/ in App Resources kopieren
# -------------------------------------------------------
echo "━━━ [2/4] Assets kopieren ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
sh "$SCRIPTS/copy-dist-to-app.sh"
echo ""

# -------------------------------------------------------
# Schritt 3: macOS App bauen
# -------------------------------------------------------
echo "━━━ [3/4] macOS App Build (xcodebuild) ━━━━━━━━━━━━━━━━━━━━━"
sh "$SCRIPTS/build-mac-app.sh"
echo ""

# -------------------------------------------------------
# Schritt 4: DMG erstellen
# -------------------------------------------------------
echo "━━━ [4/4] DMG ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
sh "$SCRIPTS/package-dmg.sh"
echo ""

# -------------------------------------------------------
# Zusammenfassung
# -------------------------------------------------------
END_TIME=$(date +%s)
ELAPSED=$(( END_TIME - START_TIME ))
DMG_PATH="$REPO_ROOT/.build/dist/InfrontToolkit-${VERSION}.dmg"

echo "╔═════════════════════════════════════���════════════════════╗"
echo "║  Release Build abgeschlossen in ${ELAPSED}s              "
echo "╚══════════════════════════════════════════���═══════════════╝"
echo ""
echo "  Version:  $VERSION"
if [ -f "$DMG_PATH" ]; then
  echo "  DMG:      $DMG_PATH"
  echo "  Größe:    $(du -sh "$DMG_PATH" | cut -f1)"
fi
echo ""
echo "  Nächste Schritte:"
echo "  • DMG testen:   open \"$DMG_PATH\""
echo "  • GitHub Release anlegen: Siehe docs/RELEASE.md"
echo ""
