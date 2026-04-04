# Infront Toolkit – Installation auf macOS

> PowerPoint Add-in + native macOS Begleit-App

---

## Voraussetzungen

| Anforderung | Version |
|-------------|---------|
| macOS | 13 Ventura oder neuer |
| Microsoft PowerPoint für Mac | Microsoft 365 (Abonnement) |
| Internetverbindung | Nur für erste Add-in-Nutzung (CDN-Icons) |

---

## Schritt 1 – App installieren

1. `InfrontToolkit-<version>.dmg` öffnen
2. **InfrontToolkit.app** in den `Applications`-Ordner ziehen
3. DMG schließen und auswerfen

> **Gatekeeper-Hinweis:** Beim ersten Start erscheint ggf. eine Warnung
> „Entwickler nicht verifiziert". Öffne dann:
> **Systemeinstellungen → Datenschutz & Sicherheit → App trotzdem öffnen**

---

## Schritt 2 – App starten & Setup-Assistent

Starte **InfrontToolkit.app** aus dem Finder oder Launchpad.
Das App-Icon erscheint als Blitz-Symbol ⚡ in der **Menu Bar** (rechts oben).

Klicke auf das Icon → **Dashboard öffnen…** — der Setup-Assistent startet automatisch beim Erststart.

### Assistent-Schritte (ca. 5 Minuten)

#### Willkommen
- Port für den lokalen Server einstellen (Standard: **3000**)
- Port nur ändern wenn 3000 bereits belegt ist
- Klick **Weiter**

#### Server starten
- Klick **Server starten** — der lokale HTTP-Server wird gestartet
- Grüne Bestätigung: „Server läuft auf Port 3000"
- Der Assistent wechselt automatisch zum nächsten Schritt

#### Manifest installieren
- Klick **Manifest installieren** — kopiert die Add-in-Beschreibungsdatei in einen lokalen Ordner
- Pfad wird angezeigt (z.B. `~/Library/Application Support/InfrontToolkit/Manifests/`)
- Klick **Weiter**

#### Office konfigurieren ← wichtigster Schritt
Dieser Schritt registriert den Manifest-Ordner in Office **einmalig**:

1. PowerPoint (oder Word/Excel) öffnen
2. Menü **Einfügen** → **Add-ins** → **Weitere Add-ins…**
3. Tab **Freigegebene Ordnerkataloge** wählen
4. Klick **Ordner in Finder öffnen** im Assistenten, dann den angezeigten Pfad kopieren
5. Pfad in das Feld „Katalog-URL" in Office einfügen → **Katalog hinzufügen**
6. Haken bei **„Zum Menü hinzufügen"** setzen → **OK**
7. Office vollständig beenden und neu starten

Nach dem Neustart erscheint der **Infront Toolkit**-Tab im PowerPoint-Ribbon.

#### Autostart (empfohlen)
- Toggle **„Bei Login automatisch starten"** aktivieren
- Stellt sicher, dass der Server immer läuft wenn PowerPoint geöffnet wird
- Klick **Weiter**

#### Fertig
- Klick **Fertig** — Setup abgeschlossen
- Die App wechselt in den normalen Betriebsmodus (Dashboard)

---

## Schritt 3 – Add-in in PowerPoint nutzen

1. PowerPoint starten (⚡ in der Menu Bar zeigt grün = Server läuft)
2. Tab **Infront Toolkit** im Ribbon wählen
3. Gewünschtes Feature aus dem Ribbon öffnen

Beim ersten Öffnen eines Features wird das Add-in vom lokalen Server geladen.
**Internet ist danach nicht mehr erforderlich** – alles läuft lokal.

---

## Normaler Betrieb

| Aktion | Beschreibung |
|--------|-------------|
| ⚡ grün (Menu Bar) | Server läuft, Add-in verfügbar |
| ⚡ durchgestrichen | Server gestoppt |
| Klick auf ⚡ → Dashboard | Server-Status, Logs, Einstellungen |
| Klick auf ⚡ → Server starten/stoppen | Server manuell steuern |

**Mit Autostart** muss die App nicht manuell gestartet werden — sie läuft im Hintergrund nach dem Login.

---

## Brand Config anpassen (optional)

Die Brand-Konfiguration (Farben, Schriften, Toleranzen) kann angepasst werden:

1. Dashboard öffnen → Tab **Einstellungen** → Abschnitt **Brand Config**
2. Klick **In Finder öffnen** — der Config-Ordner öffnet sich
3. Datei `Infront_BrandConfig.json` anlegen oder bearbeiten
4. Speichern — beim nächsten Add-in-Aufruf wird die neue Konfiguration verwendet

**Pfad:** `~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json`

Um auf die Standard-Konfiguration zurückzukehren: **Auf Standard zurücksetzen** im Einstellungs-Tab.

---

## Fehlerbehebung

### Tab „Infront Toolkit" fehlt im Ribbon

1. Prüfen ob ⚡ grün ist (Server läuft) — ggf. Dashboard → **Server starten**
2. Pfad in Office-Katalog korrekt? → Einstellungen → Manifest → **Neu installieren**
3. Office vollständig beenden (⌘+Q) und neu starten
4. Im Zweifelsfall: Schritt „Office konfigurieren" wiederholen

### Server startet nicht (Fehler: Port bereits belegt)

1. Dashboard öffnen → Einstellungen → anderen Port eintragen (z.B. 3001)
2. Manifest neu installieren (Schritt: Einstellungen → Manifest → **Neu installieren**)
3. In Office den Katalog-Pfad aktualisieren (er hat sich nicht geändert) und Office neu starten

### Add-in zeigt „Verbindung zum Server nicht möglich"

- ⚡ in der Menu Bar grün? Wenn nicht: **Server starten**
- Port im Add-in stimmt nicht mit App-Port überein → Manifest neu installieren + Office neu starten

### macOS meldet „beschädigte App" oder verweigert Start

```
Systemeinstellungen → Datenschutz & Sicherheit
→ unter "Sicherheit": "App trotzdem öffnen"
```

Oder im Terminal:
```sh
xattr -cr /Applications/InfrontToolkit.app
```

---

## Deinstallation

1. InfrontToolkit.app beenden (Menu Bar → **Beenden**)
2. `/Applications/InfrontToolkit.app` in den Papierkorb
3. Nutzerdaten entfernen (optional):
   ```sh
   rm -rf ~/Library/Application\ Support/InfrontToolkit
   ```
4. Login-Item entfernen (falls Autostart aktiviert war):
   **Systemeinstellungen → Allgemein → Anmeldeobjekte** → InfrontToolkit entfernen
5. In Office: Freigegebenen Ordnerkatalog entfernen (optional):
   **Einfügen → Add-ins → Weitere Add-ins… → Freigegebene Ordnerkataloge**

---

## Konfigurationspfade (Referenz)

| Zweck | Pfad |
|-------|------|
| Manifest | `~/Library/Application Support/InfrontToolkit/Manifests/InfrontToolkit-manifest.xml` |
| Brand Config (User) | `~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json` |
| App-Logs | Logs-Tab im Dashboard |
