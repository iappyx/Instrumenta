#!/usr/bin/env bash
# package-dmg.sh
# Erstellt eine DMG-Datei aus der gebauten macOS App.
# (Schritt 7 – noch nicht implementiert)
#
# Voraussetzungen (später):
#   - build-mac-app.sh muss erfolgreich durchgelaufen sein
#   - create-dmg (brew install create-dmg) oder hdiutil (macOS built-in)
#   - Code Signing: Apple Developer ID Application Certificate
#   - Notarization: xcrun notarytool (erfordert Apple Developer Account)
#
# WICHTIG: Ohne Code Signing + Notarization zeigt macOS Gatekeeper eine
#           Sicherheitswarnung. Für Produktion ist ein Apple Developer
#           Account ($99/Jahr) zwingend erforderlich.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

echo "==> [package-dmg] DMG-Packaging..."
echo ""
echo "    [STUB] Noch nicht implementiert (Schritt 7)."
echo ""
echo "    Geplanter Flow:"
echo "    1. xcodebuild → InfrontToolkit.app"
echo "    2. codesign --deep --sign 'Developer ID Application: ...' InfrontToolkit.app"
echo "    3. create-dmg InfrontToolkit.dmg InfrontToolkit.app"
echo "    4. xcrun notarytool submit InfrontToolkit.dmg (Apple Developer Account nötig)"
echo "    5. xcrun stapler staple InfrontToolkit.dmg"
echo ""
echo "    Siehe: docs/RELEASE.md für den vollständigen Release-Prozess."
