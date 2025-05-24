#!/bin/bash

# MIT License
# Copyright (c) 2021-2025 iappyx
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

cp ../../../../bin/InstrumentaPowerpointToolbar.ppam Instrumenta.pkg/Payload/InstrumentaPowerpointToolbar.ppam
cp ../../../../bin/InstrumentaAppleScriptPlugin.applescript Instrumenta.pkg/Payload/InstrumentaAppleScriptPlugin.applescript
cp ../../../../../Instrumenta-Keys/bin/mac/InstrumentaKeysHelper.ppam InstrumentaKeys.pkg/Payload/InstrumentaKeysHelper.ppam
cp ../../../../../Instrumenta-Keys/bin/mac/InstrumentaKeys.applescript InstrumentaKeys.pkg/Payload/InstrumentaKeys.applescript

pkgbuild --root Instrumenta.pkg/Payload --identifier com.iappyx.instrumenta --version 1.0 --install-location "/Library/Application Support/Instrumenta" --scripts Instrumenta.pkg/Scripts InstrumentaPackage.pkg
pkgbuild --root InstrumentaKeys.pkg/Payload --identifier com.iappyx.instrumentakeys --version 1.0 --install-location "/Library/Application Support/Instrumenta" --scripts InstrumentaKeys.pkg/Scripts InstrumentaKeysPackage.pkg
pkgbuild --root InstrumentaUninstaller.pkg/Payload --identifier com.iappyx.instrumenta.uninstall --version 1.0 --scripts InstrumentaUninstaller.pkg/Scripts InstrumentaUninstallerPackage.pkg

productbuild --distribution distribution.xml --resources ./Resources --package-path . InstrumentaInstaller.pkg
rm -f InstrumentaPackage.pkg
rm -f InstrumentaKeysPackage.pkg
rm -f InstrumentaUninstallerPackage.pkg

DMG_NAME="InstrumentaInstaller.dmg"
rm -f "$DMG_NAME"

VOLUME_NAME="Instrumenta Installer"
PKG_FILE="InstrumentaInstaller.pkg"
ICON_FILE="volumeIcon.icns"
BACKGROUND_FILE="background.png"
TMP_DIR="TMP"

if ! command -v create-dmg &> /dev/null; then
    echo "Installing create-dmg..."
    brew install create-dmg
fi

mkdir -p "$TMP_DIR/$VOLUME_NAME/.background"
mkdir -p "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"

cp "$PKG_FILE" "$TMP_DIR/$VOLUME_NAME/"
cp "$BACKGROUND_FILE" "$TMP_DIR/$VOLUME_NAME/.background/"
cp "../../../../bin/InstrumentaAppleScriptPlugin.applescript" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"
cp "../../../../bin/InstrumentaPowerpointToolbar.ppam" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"
cp "../../../../../Instrumenta-Keys/bin/mac/InstrumentaKeysHelper.ppam" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"
cp "../../../../../Instrumenta-Keys/bin/mac/InstrumentaKeys.applescript" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"
cp "Instrumenta_Manual_Install.applescript" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"
cp "Instrumenta_Manual_Uninstall.applescript" "$TMP_DIR/$VOLUME_NAME/Files for manual installation/"

if [ -f "$ICON_FILE" ]; then
    cp "$ICON_FILE" "$TMP_DIR/$VOLUME_NAME/.VolumeIcon.icns"
    SetFile -a V "$TMP_DIR/$VOLUME_NAME/.VolumeIcon.icns"
fi

create-dmg \
  --volname "$VOLUME_NAME" \
  --window-pos 200 200 \
  --window-size 600 400 \
  --background "background.png" \
  --hide-extension "$PKG_FILE" \
  --icon "$PKG_FILE" 150 150 \
  --icon "Files for manual installation" 350 150 \
  "$DMG_NAME" \
  "$TMP_DIR/$VOLUME_NAME/"

rm -rf "$TMP_DIR"
rm -f InstrumentaInstaller.pkg
mv InstrumentaInstaller.dmg ../../../../bin/Installers/

echo "DMG successfully created: $DMG_NAME"
