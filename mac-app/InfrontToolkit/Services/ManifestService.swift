// ManifestService.swift
// Verwaltet das Office Add-in Manifest:
// - Generiert manifest.xml aus dem Bundle-Template (ersetzt {{PORT}})
// - Installiert das Manifest in den Shared Folder Catalog Ordner
// - Dieser Ordner muss einmalig in den Office-Einstellungen registriert werden

import Foundation

// MARK: - Errors

enum ManifestError: LocalizedError {
    case templateNotFound
    case templateReadFailed(Error)
    case directoryCreationFailed(Error)
    case writeFailed(Error)

    var errorDescription: String? {
        switch self {
        case .templateNotFound:
            return "manifest-template.xml nicht im App-Bundle gefunden."
        case .templateReadFailed(let e):
            return "Template lesen fehlgeschlagen: \(e.localizedDescription)"
        case .directoryCreationFailed(let e):
            return "Katalog-Ordner erstellen fehlgeschlagen: \(e.localizedDescription)"
        case .writeFailed(let e):
            return "Manifest schreiben fehlgeschlagen: \(e.localizedDescription)"
        }
    }
}

// MARK: - ManifestService

struct ManifestService {

    // MARK: Paths

    /// Ordner, in dem das Manifest abgelegt wird.
    /// Dieser Ordner muss in Office als "Freigegebener Ordnerkatalog" registriert werden.
    static var catalogFolderURL: URL {
        let appSupport = FileManager.default
            .urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        return appSupport
            .appendingPathComponent("InfrontToolkit")
            .appendingPathComponent("Manifests")
    }

    static var installedManifestURL: URL {
        catalogFolderURL.appendingPathComponent("InfrontToolkit-manifest.xml")
    }

    // MARK: Install

    /// Generiert das Manifest aus dem Bundle-Template und installiert es im Katalog-Ordner.
    /// - Parameter port: Der Port auf dem der lokale Server läuft.
    /// - Returns: URL der installierten Manifest-Datei.
    @discardableResult
    static func installManifest(port: Int) throws -> URL {
        // 1. Template laden
        guard let templateURL = Bundle.main.url(
            forResource: "manifest-template",
            withExtension: "xml"
        ) else {
            throw ManifestError.templateNotFound
        }

        let templateContent: String
        do {
            templateContent = try String(contentsOf: templateURL, encoding: .utf8)
        } catch {
            throw ManifestError.templateReadFailed(error)
        }

        // 2. Platzhalter ersetzen
        let manifest = templateContent
            .replacingOccurrences(of: "{{PORT}}", with: String(port))

        // 3. Katalog-Ordner erstellen (falls nicht vorhanden)
        do {
            try FileManager.default.createDirectory(
                at: catalogFolderURL,
                withIntermediateDirectories: true
            )
        } catch {
            throw ManifestError.directoryCreationFailed(error)
        }

        // 4. Manifest schreiben
        do {
            try manifest.write(
                to: installedManifestURL,
                atomically: true,
                encoding: .utf8
            )
        } catch {
            throw ManifestError.writeFailed(error)
        }

        return installedManifestURL
    }

    // MARK: Status

    static var isInstalled: Bool {
        FileManager.default.fileExists(atPath: installedManifestURL.path)
    }

    static func removeManifest() throws {
        guard isInstalled else { return }
        try FileManager.default.removeItem(at: installedManifestURL)
    }
}
