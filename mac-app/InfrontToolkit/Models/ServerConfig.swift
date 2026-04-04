// ServerConfig.swift
// Konfigurationsmodell für den lokalen HTTP-Server.
// Bestimmt Port, Host und Pfade zu den gebundeten Add-in-Assets.

import Foundation

struct ServerConfig {
    let port: Int
    let host: String = "127.0.0.1"  // Nur loopback – keine externen Verbindungen

    /// URL des gebundeten addin-dist/ Verzeichnisses im App Bundle.
    /// Wird vom Build-Prozess befüllt (scripts/copy-dist-to-app.sh).
    var addinDistURL: URL {
        // Bevorzuge Bundle-Ressource (korrekt für signierte App)
        if let url = Bundle.main.url(forResource: "addin-dist", withExtension: nil) {
            return url
        }
        // Fallback: manuell aus Bundle-Pfad konstruieren
        return Bundle.main.bundleURL
            .appendingPathComponent("Contents/Resources/addin-dist")
    }

    var baseURL: URL {
        URL(string: "http://\(host):\(port)")!
    }

    /// Prüft ob addin-dist/ tatsächlich vorhanden und nicht leer ist.
    var isAddinDistPresent: Bool {
        var isDir: ObjCBool = false
        guard FileManager.default.fileExists(atPath: addinDistURL.path, isDirectory: &isDir),
              isDir.boolValue else { return false }
        let contents = try? FileManager.default.contentsOfDirectory(atPath: addinDistURL.path)
        return !(contents?.isEmpty ?? true)
    }
}
