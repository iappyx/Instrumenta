#!/usr/bin/env bash
# copy-dist-to-app.sh
# Kopiert den gebauten Add-in dist/ Inhalt in das macOS App Bundle Resources.
# Wird aufgerufen von: npm run package (in office-addin/)
# Wird aufgerufen von: build-mac-app.sh (nach xcodebuild)
#
# Ziel: mac-app/InfrontToolkit/Resources/addin-dist/
# (Schritt 4 – noch nicht vollständig implementiert)

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$REPO_ROOT/office-addin/dist"
DST="$REPO_ROOT/mac-app/InfrontToolkit/Resources/addin-dist"

echo "==> [copy-dist-to-app] Kopiere Add-in dist/ → mac-app Resources..."

if [ ! -d "$SRC" ]; then
  echo "    FEHLER: $SRC existiert nicht. Bitte zuerst npm run build ausführen."
  exit 1
fi

if [ ! -d "$DST" ]; then
  echo "    WARNUNG: Zielverzeichnis $DST existiert noch nicht."
  echo "    Wird in Schritt 3/4 angelegt, wenn mac-app/ erstellt wird."
  echo "    [STUB] Kopiervorgang übersprungen."
  exit 0
fi

rm -rf "$DST"
cp -r "$SRC" "$DST"

echo "    Kopiert: $SRC → $DST"
echo "    Inhalt:"
ls -lh "$DST/"
