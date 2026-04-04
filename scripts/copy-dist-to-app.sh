#!/usr/bin/env bash
# copy-dist-to-app.sh
# Kopiert den gebauten Add-in dist/ Inhalt in das macOS App Bundle Resources.
# Wird aufgerufen von: npm run package (in office-addin/)
# Wird aufgerufen von: build-mac-app.sh (nach xcodebuild)
#
# Ziel: mac-app/InfrontToolkit/Resources/addin-dist/
# Inkl. office-addin/config/ → addin-dist/config/ (Brand-Config für Fallback)

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$REPO_ROOT/office-addin/dist"
DST="$REPO_ROOT/mac-app/InfrontToolkit/Resources/addin-dist"
CONFIG_SRC="$REPO_ROOT/office-addin/config"

echo "==> [copy-dist-to-app] Kopiere Add-in dist/ → mac-app Resources..."

if [ ! -d "$SRC" ]; then
  echo "    FEHLER: $SRC existiert nicht. Bitte zuerst npm run build ausführen."
  exit 1
fi

mkdir -p "$DST"
rm -rf "$DST"
cp -r "$SRC" "$DST"

echo "    Kopiert: $SRC → $DST"

# Brand-Config mitliefern (wird von ConfigSyncService als Bundle-Default verwendet)
if [ -d "$CONFIG_SRC" ]; then
  mkdir -p "$DST/config"
  cp -r "$CONFIG_SRC/." "$DST/config/"
  echo "    Kopiert: $CONFIG_SRC → $DST/config/"
else
  echo "    INFO: $CONFIG_SRC nicht gefunden – kein Config-Ordner kopiert."
fi

echo "    Inhalt:"
ls -lh "$DST/"
