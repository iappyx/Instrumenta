// AppAutoStartService.swift
// Registriert / deregistriert die App als macOS Login Item.
// Verwendet SMAppService (macOS 13+) – kein deprecated LaunchServices.
//
// Hinweis: Die App muss für dieses Feature korrekt signiert sein.
// Im Development-Build (unsigniert) schlägt die Registrierung möglicherweise fehl.

import Foundation
import ServiceManagement

// MARK: - Errors

enum AutoStartError: LocalizedError {
    case registrationFailed(Error)
    case unregistrationFailed(Error)

    var errorDescription: String? {
        switch self {
        case .registrationFailed(let e):
            return "Autostart aktivieren fehlgeschlagen: \(e.localizedDescription)"
        case .unregistrationFailed(let e):
            return "Autostart deaktivieren fehlgeschlagen: \(e.localizedDescription)"
        }
    }
}

// MARK: - AppAutoStartService

struct AppAutoStartService {

    private static let service = SMAppService.mainApp

    /// Gibt an, ob die App als Login Item registriert ist.
    static var isEnabled: Bool {
        service.status == .enabled
    }

    /// Registriert die App als Login Item.
    /// Wirft AutoStartError.registrationFailed bei Fehler.
    static func enable() throws {
        do {
            try service.register()
        } catch {
            throw AutoStartError.registrationFailed(error)
        }
    }

    /// Deregistriert die App als Login Item.
    /// Wirft AutoStartError.unregistrationFailed bei Fehler.
    static func disable() throws {
        do {
            try service.unregister()
        } catch {
            throw AutoStartError.unregistrationFailed(error)
        }
    }
}
