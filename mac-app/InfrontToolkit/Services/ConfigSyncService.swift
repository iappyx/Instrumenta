// ConfigSyncService.swift
// Verwaltet die Brand-Konfiguration (Infront_BrandConfig.json) auf der macOS-Seite.
//
// Priorität beim Laden:
//   1. User-Override: ~/Library/Application Support/InfrontToolkit/config/Infront_BrandConfig.json
//   2. Bundle-Default: app bundle Resources/Infront_BrandConfig.json
//   3. Eingebettetes Fallback (entspricht DEFAULT_BRAND_CONFIG im Add-in)
//
// Diese Datei wird vom LocalServerService unter /config/Infront_BrandConfig.json
// ausgeliefert, sodass das Add-in die Konfiguration per HTTP lesen kann.

import Foundation

// MARK: - Errors

enum ConfigSyncError: LocalizedError {
    case readFailed(Error)
    case writeFailed(Error)
    case invalidJSON

    var errorDescription: String? {
        switch self {
        case .readFailed(let e):  return "Konfiguration lesen fehlgeschlagen: \(e.localizedDescription)"
        case .writeFailed(let e): return "Konfiguration schreiben fehlgeschlagen: \(e.localizedDescription)"
        case .invalidJSON:        return "Konfigurationsdatei enthält ungültiges JSON."
        }
    }
}

// MARK: - ConfigSyncService

struct ConfigSyncService {

    // MARK: Paths

    /// User-Override: Hier kann der Nutzer die Konfiguration anpassen.
    static var userConfigURL: URL {
        let appSupport = FileManager.default
            .urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        return appSupport
            .appendingPathComponent("InfrontToolkit")
            .appendingPathComponent("config")
            .appendingPathComponent("Infront_BrandConfig.json")
    }

    /// Bundle-Default: Mitgelieferte Standardkonfiguration.
    static var bundleConfigURL: URL? {
        Bundle.main.url(forResource: "Infront_BrandConfig", withExtension: "json")
    }

    // MARK: Read

    /// Gibt die aktuelle Konfiguration als JSON-Data zurück.
    /// Reihenfolge: User-Override → Bundle-Default → Eingebettetes Fallback.
    static func currentConfigData() -> Data {
        // 1. User-Override
        if FileManager.default.fileExists(atPath: userConfigURL.path),
           let data = try? Data(contentsOf: userConfigURL),
           isValidJSON(data) {
            return data
        }

        // 2. Bundle-Default
        if let bundleURL = bundleConfigURL,
           let data = try? Data(contentsOf: bundleURL),
           isValidJSON(data) {
            return data
        }

        // 3. Eingebettetes Fallback (passt zu DEFAULT_BRAND_CONFIG im Add-in)
        return embeddedDefaultConfig
    }

    /// Gibt die aktuelle Konfiguration als hübsch formatierten JSON-String zurück.
    static func currentConfigJSON() -> String {
        let data = currentConfigData()
        if let json = try? JSONSerialization.jsonObject(with: data),
           let pretty = try? JSONSerialization.data(withJSONObject: json, options: .prettyPrinted),
           let str = String(data: pretty, encoding: .utf8) {
            return str
        }
        return String(data: data, encoding: .utf8) ?? "{}"
    }

    // MARK: Write

    /// Schreibt eine neue Konfiguration als User-Override.
    /// - Parameter jsonData: Gültiger JSON-Data der Brand-Konfiguration.
    static func saveUserConfig(_ jsonData: Data) throws {
        guard isValidJSON(jsonData) else { throw ConfigSyncError.invalidJSON }

        do {
            try FileManager.default.createDirectory(
                at: userConfigURL.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try jsonData.write(to: userConfigURL, options: .atomic)
        } catch {
            throw ConfigSyncError.writeFailed(error)
        }
    }

    /// Entfernt den User-Override (nächste Anfrage erhält wieder den Bundle-Default).
    static func resetToDefault() throws {
        guard FileManager.default.fileExists(atPath: userConfigURL.path) else { return }
        do {
            try FileManager.default.removeItem(at: userConfigURL)
        } catch {
            throw ConfigSyncError.writeFailed(error)
        }
    }

    static var hasUserOverride: Bool {
        FileManager.default.fileExists(atPath: userConfigURL.path)
    }

    // MARK: Helpers

    private static func isValidJSON(_ data: Data) -> Bool {
        (try? JSONSerialization.jsonObject(with: data)) != nil
    }

    // MARK: Embedded fallback (entspricht DEFAULT_BRAND_CONFIG im TypeScript Add-in)

    private static let embeddedDefaultConfig: Data = {
        let json = """
        {
          "version": "1.0",
          "activeProfile": "Default",
          "profiles": [
            {
              "name": "Default",
              "allowedFonts": ["Calibri", "Arial", "Helvetica Neue"],
              "brandColors": [
                { "name": "Infront Navy",  "value": "#003366" },
                { "name": "Infront Red",   "value": "#FF0000" },
                { "name": "Weiß",          "value": "#FFFFFF" },
                { "name": "Schwarz",       "value": "#000000" },
                { "name": "Hellgrau",      "value": "#CCCCCC" }
              ],
              "colorTolerance": 10,
              "minFontSizePt": 8
            },
            {
              "name": "Strict",
              "allowedFonts": ["Calibri"],
              "brandColors": [
                { "name": "Infront Navy",  "value": "#003366" },
                { "name": "Weiß",          "value": "#FFFFFF" },
                { "name": "Schwarz",       "value": "#000000" }
              ],
              "colorTolerance": 0,
              "minFontSizePt": 10
            }
          ]
        }
        """
        return Data(json.utf8)
    }()
}
