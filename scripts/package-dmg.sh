#!/usr/bin/env bash
# package-dmg.sh
# Erstellt eine verteilbare DMG-Datei aus der gebauten InfrontToolkit.app.
#
# Voraussetzungen:
#   - build-mac-app.sh muss vorher erfolgreich gelaufen sein
#   - hdiutil (macOS built-in) – immer verfügbar
#   - create-dmg (optional, brew install create-dmg) – für hübscheres DMG mit Hintergrund
#
# Code Signing / Notarization (Produktion):
#   - Apple Developer ID Application Certificate im Keychain
#   - CODESIGN_IDENTITY: z.B. "Developer ID Application: Infront AG (XXXXXXXXXX)"
#   - NOTARIZATION_PROFILE: xcrun notarytool store-credentials (einmalig einrichten)
#   - Für Entwicklung: SKIP_SIGNING=1 überspringt Signierung + Notarization
#
# Umgebungsvariablen:
#   CODESIGN_IDENTITY  – Code-Signing-Zertifikat (optional, für Produktion)
#   NOTARIZATION_PROFILE – notarytool keychain profile name (optional)
#   SKIP_SIGNING=1     – Signierung überspringen (Entwicklung/CI ohne Zertifikate)
#
# Ausgabe:
#   .build/dist/InfrontToolkit-<version>.dmg

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="$(cat "$REPO_ROOT/VERSION" | tr -d '[:space:]')"
BUILD_DIR="$REPO_ROOT/.build"
APP_PATH="$BUILD_DIR/mac-app/Build/Products/Release/InfrontToolkit.app"
DIST_DIR="$BUILD_DIR/dist"
DMG_NAME="InfrontToolkit-${VERSION}.dmg"
DMG_PATH="$DIST_DIR/$DMG_NAME"
VOLUME_NAME="Infront Toolkit $VERSION"
SKIP_SIGNING="${SKIP_SIGNING:-0}"

echo "==> [package-dmg] Infront Toolkit $VERSION"
echo "    App:    $APP_PATH"
echo "    Output: $DMG_PATH"
echo ""

# -------------------------------------------------------
# 0. Voraussetzungen prüfen
# -------------------------------------------------------

# App gefunden?
if [ ! -d "$APP_PATH" ]; then
  # Alternativer Pfad (falls xcodebuild eine andere Struktur erzeugt hat)
  FALLBACK=$(find "$BUILD_DIR" -name "InfrontToolkit.app" -type d 2>/dev/null | head -1 || true)
  if [ -n "$FALLBACK" ]; then
    APP_PATH="$FALLBACK"
    echo "    INFO: App gefunden unter: $APP_PATH"
  else
    echo "    FEHLER: InfrontToolkit.app nicht gefunden."
    echo "    Bitte zuerst ausführen: sh scripts/build-mac-app.sh"
    exit 1
  fi
fi

mkdir -p "$DIST_DIR"

# -------------------------------------------------------
# 1. Code Signing (optional)
# -------------------------------------------------------

if [ "$SKIP_SIGNING" = "1" ]; then
  echo "==> [1/4] Code Signing übersprungen (SKIP_SIGNING=1)"
elif [ -n "${CODESIGN_IDENTITY:-}" ]; then
  echo "==> [1/4] Code Signing mit: $CODESIGN_IDENTITY"
  codesign \
    --deep \
    --force \
    --options runtime \
    --sign "$CODESIGN_IDENTITY" \
    "$APP_PATH"
  echo "    OK: App signiert"

  # Signierung verifizieren
  codesign --verify --verbose "$APP_PATH" 2>&1 | head -3 || true
else
  echo "==> [1/4] WARNUNG: CODESIGN_IDENTITY nicht gesetzt."
  echo "    Die App wird NICHT signiert. Gatekeeper zeigt eine Sicherheitswarnung."
  echo "    Für Produktion: CODESIGN_IDENTITY=\"Developer ID Application: ...\" setzen."
  echo "    Für Entwicklung: SKIP_SIGNING=1 setzen um diese Warnung zu unterdrücken."
fi

# -------------------------------------------------------
# 2. DMG erstellen
# -------------------------------------------------------

echo "==> [2/4] DMG erstellen..."

if command -v create-dmg &>/dev/null && [ "$SKIP_SIGNING" != "1" -o -n "${CODESIGN_IDENTITY:-}" ]; then
  echo "    Methode: create-dmg (hübsches DMG mit App-Fenster)"
  create-dmg \
    --volname "$VOLUME_NAME" \
    --window-pos 200 120 \
    --window-size 540 360 \
    --icon-size 120 \
    --icon "InfrontToolkit.app" 160 170 \
    --hide-extension "InfrontToolkit.app" \
    --app-drop-link 380 170 \
    --no-internet-enable \
    "$DMG_PATH" \
    "$APP_PATH" 2>&1 | tail -5 || true
else
  echo "    Methode: hdiutil (Standard macOS Tool)"
  _build_dmg_hdiutil
fi

function _build_dmg_hdiutil() {
  local TMP_DMG="$DIST_DIR/InfrontToolkit-tmp-$$.dmg"
  local MOUNT_POINT="/Volumes/$VOLUME_NAME"

  # Cleanup bei Abbruch
  trap "hdiutil detach '$MOUNT_POINT' 2>/dev/null || true; rm -f '$TMP_DMG'" EXIT

  # Lese-/Schreib-DMG anlegen (etwas größer als die App)
  local APP_SIZE_KB
  APP_SIZE_KB=$(du -sk "$APP_PATH" | cut -f1)
  local DMG_SIZE_MB=$(( (APP_SIZE_KB / 1024) + 20 ))

  hdiutil create \
    -size "${DMG_SIZE_MB}m" \
    -fs HFS+ \
    -volname "$VOLUME_NAME" \
    -ov \
    -o "$TMP_DMG" \
    > /dev/null

  # Mounten
  hdiutil attach "$TMP_DMG" -mountpoint "$MOUNT_POINT" -nobrowse -quiet

  # App und /Applications-Symlink kopieren
  cp -r "$APP_PATH" "$MOUNT_POINT/"
  ln -s /Applications "$MOUNT_POINT/Applications"

  # Unmounten
  hdiutil detach "$MOUNT_POINT" -quiet

  # In komprimiertes, schreibgeschütztes DMG konvertieren
  hdiutil convert "$TMP_DMG" \
    -format UDZO \
    -imagekey zlib-level=9 \
    -ov \
    -o "$DMG_PATH" \
    > /dev/null

  rm -f "$TMP_DMG"
  trap - EXIT
}

# hdiutil aufrufen wenn create-dmg nicht verfügbar
if ! command -v create-dmg &>/dev/null; then
  _build_dmg_hdiutil
fi

if [ ! -f "$DMG_PATH" ]; then
  echo "    FEHLER: DMG wurde nicht erstellt: $DMG_PATH"
  exit 1
fi

echo "    OK: $DMG_PATH ($(du -sh "$DMG_PATH" | cut -f1))"

# -------------------------------------------------------
# 3. Notarization (optional)
# -------------------------------------------------------

if [ -n "${NOTARIZATION_PROFILE:-}" ] && [ "$SKIP_SIGNING" != "1" ]; then
  echo "==> [3/4] Notarization wird eingereicht..."
  xcrun notarytool submit "$DMG_PATH" \
    --keychain-profile "$NOTARIZATION_PROFILE" \
    --wait
  echo "    Stapler läuft..."
  xcrun stapler staple "$DMG_PATH"
  echo "    OK: DMG notarisiert + gestapelt"
elif [ "$SKIP_SIGNING" = "1" ]; then
  echo "==> [3/4] Notarization übersprungen (SKIP_SIGNING=1)"
else
  echo "==> [3/4] Notarization übersprungen (NOTARIZATION_PROFILE nicht gesetzt)"
  echo "    Für Produktion: xcrun notarytool store-credentials --apple-id ... --team-id ..."
fi

# -------------------------------------------------------
# 4. Zusammenfassung
# -------------------------------------------------------

echo ""
echo "==> [4/4] DMG fertig."
echo "    Datei:   $DMG_PATH"
echo "    Version: $VERSION"
echo "    Größe:   $(du -sh "$DMG_PATH" | cut -f1)"
echo ""

if [ "$SKIP_SIGNING" = "1" ] || [ -z "${CODESIGN_IDENTITY:-}" ]; then
  echo "HINWEIS: Die App ist NICHT signiert/notarisiert."
  echo "         Nutzer sehen beim ersten Start eine Gatekeeper-Warnung."
  echo "         Workaround: xattr -cr /Applications/InfrontToolkit.app"
  echo "         Für Produktion: CODESIGN_IDENTITY + NOTARIZATION_PROFILE setzen."
fi

echo ""
echo "Nächste Schritte:"
echo "  • Testen:   open \"$DMG_PATH\""
echo "  • Release:  Siehe docs/RELEASE.md"
