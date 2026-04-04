#!/usr/bin/env bash
# build-addin.sh
# Baut das Office.js Add-in (Production Build → office-addin/dist/).
# Ausgabe: office-addin/dist/ mit taskpane.html, commands.html, taskpane.js,
#           commands.js und assets/

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ADDIN_DIR="$REPO_ROOT/office-addin"

echo "==> [build-addin] Starte Production Build..."
echo "    Verzeichnis: $ADDIN_DIR"

cd "$ADDIN_DIR"

if [ ! -d node_modules ]; then
  echo "    node_modules fehlt – führe npm install aus..."
  npm install
fi

npm run build

echo ""
echo "==> Build abgeschlossen."
echo "    Output: $ADDIN_DIR/dist/"
ls -lh "$ADDIN_DIR/dist/" 2>/dev/null || echo "    (dist/ noch nicht vorhanden)"
