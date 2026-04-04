#!/usr/bin/env bash
# build-mac-app.sh
# Baut die native macOS Begleit-App (Schritt 3 – noch nicht implementiert).
# Dieses Skript wird in Schritt 3 vollständig ausgebaut.
#
# Voraussetzungen (später):
#   - Xcode 15+
#   - macOS 13+ SDK
#   - Apple Developer Account für Code Signing (optional, Produktion)

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
MAC_APP_DIR="$REPO_ROOT/mac-app"

echo "==> [build-mac-app] Prüfe Voraussetzungen..."

if [ ! -d "$MAC_APP_DIR" ]; then
  echo "    FEHLER: mac-app/ Verzeichnis existiert noch nicht."
  echo "    → Wird in Schritt 3 angelegt."
  exit 1
fi

# STUB: Wird in Schritt 3 implementiert
# xcodebuild -project "$MAC_APP_DIR/InfrontToolkit.xcodeproj" \
#   -scheme InfrontToolkit \
#   -configuration Release \
#   -derivedDataPath "$REPO_ROOT/.build/mac-app" \
#   clean build

echo ""
echo "    [STUB] mac-app Build noch nicht implementiert (Schritt 3)."
echo "    Siehe: docs/ARCHITECTURE.md für den geplanten Build-Flow."
