// AppState.swift
// Zentraler Observable-State der App.
// Wird per @EnvironmentObject in alle Views injiziert.

import Foundation
import SwiftUI

// MARK: - ServerStatus

enum ServerStatus: Equatable {
    case stopped
    case starting
    case running(port: Int)
    case error(String)

    var isRunning: Bool {
        if case .running = self { return true }
        return false
    }

    var description: String {
        switch self {
        case .stopped:            return "Server gestoppt"
        case .starting:           return "Server startet…"
        case .running(let port):  return "Läuft · Port \(port)"
        case .error(let msg):     return "Fehler: \(msg)"
        }
    }

    var statusColor: Color {
        switch self {
        case .stopped:  return .secondary
        case .starting: return .orange
        case .running:  return .green
        case .error:    return .red
        }
    }

    var menuBarIconName: String {
        switch self {
        case .running: return "bolt.fill"
        case .error:   return "bolt.slash.fill"
        default:       return "bolt.slash"
        }
    }
}

// MARK: - SetupStep

enum SetupStep: Int, CaseIterable {
    case welcome      = 0
    case startServer  = 1
    case manifest     = 2
    case officeCatalog = 3
    case autostart    = 4
    case done         = 5

    var title: String {
        switch self {
        case .welcome:       return "Willkommen"
        case .startServer:   return "Server starten"
        case .manifest:      return "Manifest installieren"
        case .officeCatalog: return "Office konfigurieren"
        case .autostart:     return "Autostart"
        case .done:          return "Fertig"
        }
    }
}

// MARK: - LogEntry

struct LogEntry: Identifiable {
    let id = UUID()
    let timestamp: Date
    let level: Level
    let message: String

    enum Level: String {
        case info    = "INFO"
        case warning = "WARN"
        case error   = "ERR "

        var color: Color {
            switch self {
            case .info:    return .secondary
            case .warning: return .orange
            case .error:   return .red
            }
        }
    }

    var formattedTimestamp: String {
        let fmt = DateFormatter()
        fmt.dateFormat = "HH:mm:ss"
        return fmt.string(from: timestamp)
    }
}

// MARK: - AppState

@MainActor
final class AppState: ObservableObject {

    // MARK: Published state

    @Published var serverStatus: ServerStatus = .stopped
    @Published var setupCompleted: Bool = false
    @Published var currentSetupStep: SetupStep = .welcome
    @Published var selectedPort: Int = 3000
    @Published var autoStartEnabled: Bool = false
    @Published var logEntries: [LogEntry] = []
    @Published var manifestInstalledPath: URL?

    // MARK: Services

    private let serverService = LocalServerService()

    // MARK: UserDefaults keys

    private enum Keys {
        static let setupCompleted = "setupCompleted"
        static let selectedPort   = "selectedPort"
        static let autoStart      = "autoStartEnabled"
    }

    // MARK: Init

    init() {
        let defaults = UserDefaults.standard
        setupCompleted   = defaults.bool(forKey: Keys.setupCompleted)
        let port         = defaults.integer(forKey: Keys.selectedPort)
        selectedPort     = port > 1024 ? port : 3000
        autoStartEnabled = defaults.bool(forKey: Keys.autoStart)

        // Manifest-Pfad wiederherstellen wenn bereits installiert (z.B. nach Neustart)
        if ManifestService.isInstalled {
            manifestInstalledPath = ManifestService.installedManifestURL
        }

        // Auto-Start: Server beim App-Launch starten wenn Setup bereits abgeschlossen
        Task { @MainActor [weak self] in
            await self?.startIfNeeded()
        }
    }

    // MARK: Server Control

    /// Startet den Server automatisch beim App-Launch, wenn Setup abgeschlossen ist.
    func startIfNeeded() async {
        guard setupCompleted, !serverStatus.isRunning else { return }
        await startServer()
    }

    func startServer() async {
        guard !serverStatus.isRunning else { return }
        serverStatus = .starting
        addLog(.info, "Starte lokalen Server auf Port \(selectedPort)…")

        do {
            try await serverService.start(port: UInt16(selectedPort))
            await setupAddinLogHandler()
            serverStatus = .running(port: selectedPort)
            addLog(.info, "Server läuft auf http://localhost:\(selectedPort)")
        } catch {
            serverStatus = .error(error.localizedDescription)
            addLog(.error, "Server-Start fehlgeschlagen: \(error.localizedDescription)")
        }
    }

    private func setupAddinLogHandler() async {
        await serverService.setLogHandler { [weak self] level, message in
            Task { @MainActor [weak self] in
                self?.addLog(level, "[Add-in] \(message)")
            }
        }
    }

    func stopServer() async {
        await serverService.stop()
        serverStatus = .stopped
        addLog(.info, "Server gestoppt.")
    }

    // MARK: Setup

    func advanceSetupStep() {
        let next = currentSetupStep.rawValue + 1
        if let step = SetupStep(rawValue: next) {
            currentSetupStep = step
        }
    }

    func retreatSetupStep() {
        let prev = currentSetupStep.rawValue - 1
        if let step = SetupStep(rawValue: prev) {
            currentSetupStep = step
        }
    }

    func completeSetup() {
        setupCompleted = true
        UserDefaults.standard.set(true, forKey: Keys.setupCompleted)
        UserDefaults.standard.set(selectedPort, forKey: Keys.selectedPort)
        addLog(.info, "Setup abgeschlossen.")
    }

    func savePort(_ port: Int) {
        guard port > 1024, port < 65535 else { return }
        selectedPort = port
        UserDefaults.standard.set(port, forKey: Keys.selectedPort)
    }

    func saveAutoStart(_ enabled: Bool) {
        autoStartEnabled = enabled
        UserDefaults.standard.set(enabled, forKey: Keys.autoStart)
    }

    // MARK: Logging

    func addLog(_ level: LogEntry.Level, _ message: String) {
        let entry = LogEntry(timestamp: Date(), level: level, message: message)
        logEntries.append(entry)
        // Maximale Log-Größe: 500 Einträge
        if logEntries.count > 500 {
            logEntries.removeFirst(logEntries.count - 500)
        }
    }
}
