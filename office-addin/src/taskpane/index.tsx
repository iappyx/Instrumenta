/**
 * index.tsx – React-Einstiegspunkt für die Infront Toolkit Task Pane.
 * Wartet auf Office.onReady() bevor React gemountet wird.
 *
 * macOS Begleit-App Integration (optional, fällt bei Fehler auf Default zurück):
 * - Remote Logging: Alle logger.info/warn/error Aufrufe werden zusätzlich an die
 *   macOS App gesendet und erscheinen dort im Logs-Tab.
 * - Brand Config: Aktuelle Konfiguration wird beim Start vom lokalen Server geladen
 *   und in den Document Settings gespeichert (überschreibt ggf. alte Version).
 */

import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./App";
import { enableRemoteLogging } from "../utils/logger";
import { getMacAppBaseUrl, fetchBrandConfigFromServer, saveBrandConfig } from "../services/config/ConfigService";

/* global Office */

Office.onReady(() => {
  // macOS Begleit-App: Remote Logging aktivieren (fire-and-forget)
  const macAppUrl = getMacAppBaseUrl();
  enableRemoteLogging(macAppUrl);

  // macOS Begleit-App: Brand Config im Hintergrund aktualisieren
  fetchBrandConfigFromServer()
    .then((config) => {
      if (config) {
        return saveBrandConfig(config);
      }
    })
    .catch(() => {
      // Fehler sind erwartet wenn die macOS App nicht läuft – kein Problem
    });

  ReactDOM.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
    document.getElementById("root")
  );
});
