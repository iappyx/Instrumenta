#!/usr/bin/env bash
# setup-dev.sh
# Richtet die Entwicklungsumgebung für das Infront Toolkit ein.
# Voraussetzungen: Node.js ≥ 18, npm ≥ 9, Xcode Command Line Tools (für mac-app)

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ADDIN_DIR="$REPO_ROOT/office-addin"

echo "==> Infront Toolkit: Entwicklungsumgebung einrichten"
echo "    Repository: $REPO_ROOT"
echo ""

# -------------------------------------------------------
# 1. Node.js Add-in Abhängigkeiten installieren
# -------------------------------------------------------
echo "==> [1/2] npm install in office-addin/..."
cd "$ADDIN_DIR"
npm install
echo "    OK"

# -------------------------------------------------------
# 2. assets/icons/ Platzhalter prüfen
# -------------------------------------------------------
ICONS_DIR="$ADDIN_DIR/assets/icons"
for SIZE in 16 32 80; do
  ICON="$ICONS_DIR/icon-${SIZE}.png"
  if [ ! -f "$ICON" ]; then
    echo "    WARNUNG: $ICON fehlt – bitte manuell erstellen."
  fi
done

echo ""
echo "==> Setup abgeschlossen."
echo ""
echo "Nächste Schritte:"
echo "  1. Dev-Server starten:  cd office-addin && npm run dev-server"
echo "  2. PowerPoint öffnen und manifest.xml sideloaden:"
echo "     Einfügen → Add-ins → Meine Add-ins → Add-in hochladen"
echo "     Datei: office-addin/manifest.xml"
echo ""
echo "Siehe: docs/INSTALL_MAC.md für den vollständigen Installationsfluss."
