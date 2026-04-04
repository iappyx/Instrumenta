# Infront Toolkit – Integrations-Dokumentation

> Office.js Add-in ↔ native macOS Begleit-App

---

## Architektur-Übersicht

```
PowerPoint (WKWebView)              macOS App (NWListener, 127.0.0.1:PORT)
      │                                           │
      │  GET /health          ──────────────────▶ │  {"status":"ok","version":"..."}
      │  GET /config/Brand..  ──────────────────▶ │  BrandConfig JSON (User oder Default)
      │  POST /log            ──────────────────▶ │  Log-Eintrag im Logs-Tab
      │                                           │
      │  GET /taskpane.html   ──────────────────▶ │  Add-in HTML (aus addin-dist/)
      │  GET /commands.html   ──────────────────▶ │  Commands HTML
      │  GET /taskpane.js     ──────────────────▶ │  React Bundle
      │  GET /assets/...      ──────────────────▶ │  Icons
      │                                           │
```

**Kommunikationsrichtung:** Immer Add-in → macOS App (HTTP Pull/POST). Niemals push von macOS App → Add-in.

---

## HTTP-Endpunkte (LocalServerService)

### `GET /health`

Gibt den Status des lokalen Servers zurück.

**Response:**
```json
HTTP/1.1 200 OK
Content-Type: application/json

{
  "status": "ok",
  "version": "1.1.0",
  "addinDistPresent": true
}
```

**Verwendung:**
- macOS App prüft periodisch ob der Server läuft
- Add-in kann `enableRemoteLogging()` oder `fetchBrandConfigFromServer()` erst nach positiver Health-Antwort aktivieren

---

### `GET /config/Infront_BrandConfig.json`

Liefert die aktuelle Brand-Konfiguration.

**Priorität (ConfigSyncService):**
1. User-Override: `~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json`
2. Bundle-Default: gebundelte Standardkonfiguration
3. Eingebettetes Fallback: entspricht `DEFAULT_BRAND_CONFIG` im TypeScript-Add-in

**Response:**
```json
HTTP/1.1 200 OK
Content-Type: application/json

{
  "version": "1.0",
  "activeProfile": "Default",
  "profiles": [...]
}
```

**Verwendung im Add-in (`ConfigService.ts`):**
```typescript
// Optionaler Fetch – fällt bei Fehler auf DEFAULT_BRAND_CONFIG zurück
const config = await fetchBrandConfigFromServer();
if (config) {
  // In Document Settings speichern für offline-Nutzung
  await saveBrandConfig(config);
}
```

---

### `POST /log`

Empfängt Log-Einträge vom Add-in und zeigt sie im macOS App Logs-Tab.

**Request Body:**
```json
{
  "level": "info",    // "debug" | "info" | "warn" | "error"
  "module": "BrandCheckService",
  "message": "Brand Check gestartet",
  "timestamp": 1714825000000
}
```

**Response:** `HTTP/1.1 200 OK` (leer)

**Verwendung im Add-in (`logger.ts`):**
```typescript
// Remote Logging aktivieren (nach Health-Check)
enableRemoteLogging("http://localhost:3000");

// Danach werden alle log() Aufrufe zusätzlich an die macOS App gesendet
logger.info("BrandCheck", "Prüfung gestartet");
// → console.info + fire-and-forget POST /log
```

**Fire-and-forget:** Das Add-in wartet NICHT auf die Antwort. Fehler werden still ignoriert.

---

### `GET /` (statische Dateien)

Alle anderen GET-Anfragen werden als statische Datei-Anfragen behandelt.

- Pfad wird gegen `addin-dist/` aufgelöst
- Path-Traversal-Prävention: `..` Komponenten werden gefiltert
- MIME-Types: HTML, JS, CSS, JSON, PNG, SVG, WOFF2 u.a.
- CORS: `Access-Control-Allow-Origin: *` (loopback-only, kein Sicherheitsproblem)

---

## Was technisch NICHT möglich ist

| Feature | Grund | Alternative |
|---------|-------|-------------|
| Push von macOS App → Add-in | Office kontrolliert die WKWebView; kein Zugriff auf `window` von außen | Add-in pollt `/health` (nicht implementiert: zu fragil) |
| `window.webkit.messageHandlers` | Office's WKWebView ist sandboxed; die App kann keine Message Handler registrieren | HTTP-basierte Kommunikation (implementiert) |
| WebSocket-Bridge | Würde HTTPS erfordern; zu komplex für den Mehrwert | Einfache HTTP-Endpoints (implementiert) |
| Automatische Config-Updates im Add-in | Add-in hat keinen Mechanismus für Server-Push | "Config laden"-Button im Add-in (Schritt 5) |
| GlobalNotification / NSDistributedNotification | Nicht von WebView aus zugänglich | Nicht nötig |

---

## Integration aktivieren (Add-in)

Die Integration ist **opt-in** und fällt bei Nichtverfügbarkeit immer auf den lokalen Fallback zurück.

### Config vom Server laden

```typescript
import { fetchBrandConfigFromServer, saveBrandConfig } from "../services/config/ConfigService";

// In einem Feature-Panel (z.B. BrandCheckPanel):
async function loadConfigFromMacApp() {
  const config = await fetchBrandConfigFromServer();
  if (config) {
    await saveBrandConfig(config);  // in Document Settings speichern
    showNotification("Konfiguration vom Infront Toolkit geladen.");
  } else {
    showNotification("macOS App nicht erreichbar – lokale Konfiguration wird verwendet.");
  }
}
```

### Remote Logging aktivieren

```typescript
import { enableRemoteLogging } from "../utils/logger";

// In index.tsx nach Office.onReady():
Office.onReady(() => {
  // Nur aktivieren wenn macOS App erreichbar (optional: Health-Check vorher)
  enableRemoteLogging("http://localhost:3000");
  // ...
});
```

---

## Sicherheitsbetrachtung

| Aspekt | Maßnahme |
|--------|---------|
| Server nur auf loopback | `NWListener` lauscht auf `127.0.0.1`, nicht auf `0.0.0.0` |
| Path-Traversal | `..` und leere Pfadkomponenten werden in `resolveFileURL` gefiltert |
| POST-Body-Größe | Maximum 16 KB (NWListener `maximumLength`) |
| Keine Auth nötig | Nur loopback-Zugriff; andere Prozesse auf dem Mac könnten theoretisch posten – akzeptiertes Risiko für lokale Entwickler-Tools |
| CORS `*` | Nur auf loopback; kein Angriffsszenario von externen Origins |

---

## Konfigurationspfade (macOS)

| Zweck | Pfad |
|-------|------|
| Brand Config (User-Override) | `~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json` |
| Manifest (installiert) | `~/Library/Application Support/InfrontToolkit/Manifests/InfrontToolkit-manifest.xml` |
| Add-in Assets (gebundelt) | `InfrontToolkit.app/Contents/Resources/addin-dist/` |
| Log-Datei (zukünftig) | `~/Library/Logs/InfrontToolkit/infront-toolkit.log` |
