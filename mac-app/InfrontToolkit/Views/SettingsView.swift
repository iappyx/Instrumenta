// SettingsView.swift
// Einstellungen: Port, Autostart, Manifest-Verwaltung.

import SwiftUI

struct SettingsView: View {

    @EnvironmentObject var appState: AppState
    @State private var portInput: String = ""
    @State private var errorMessage: String?

    var body: some View {
        Form {
            serverSection
            autoStartSection
            manifestSection
            brandConfigSection

            if let error = errorMessage {
                Section {
                    Label(error, systemImage: "exclamationmark.triangle")
                        .foregroundColor(.red)
                        .font(.caption)
                }
            }
        }
        .formStyle(.grouped)
        .onAppear { portInput = String(appState.selectedPort) }
    }

    // MARK: - Sections

    private var serverSection: some View {
        Section("Server") {
            HStack {
                Text("Port")
                Spacer()
                TextField("3000", text: $portInput)
                    .textFieldStyle(.roundedBorder)
                    .frame(width: 80)
                    .multilineTextAlignment(.trailing)
                    .onChange(of: portInput) { val in
                        if let p = Int(val) { appState.savePort(p) }
                    }
            }
            Text("Standard: 3000. Port-Änderungen erfordern einen Server-Neustart.")
                .font(.caption)
                .foregroundColor(.secondary)

            if appState.serverStatus.isRunning {
                HStack {
                    Button("Server neu starten") {
                        Task {
                            await appState.stopServer()
                            await appState.startServer()
                        }
                    }
                    .buttonStyle(.bordered)
                    Spacer()
                }
            }
        }
    }

    private var autoStartSection: some View {
        Section("Autostart") {
            Toggle(
                "Bei macOS-Login automatisch starten",
                isOn: Binding(
                    get: { appState.autoStartEnabled },
                    set: { enabled in
                        errorMessage = nil
                        do {
                            if enabled { try AppAutoStartService.enable() }
                            else       { try AppAutoStartService.disable() }
                            appState.saveAutoStart(enabled)
                        } catch {
                            errorMessage = error.localizedDescription
                        }
                    }
                )
            )
            Text("Empfohlen, damit der Server immer verfügbar ist wenn PowerPoint geöffnet wird.")
                .font(.caption)
                .foregroundColor(.secondary)
        }
    }

    private var brandConfigSection: some View {
        Section("Brand Config") {
            HStack {
                VStack(alignment: .leading, spacing: 4) {
                    Label(
                        ConfigSyncService.hasUserOverride ? "User-Override aktiv" : "Bundle-Standard wird verwendet",
                        systemImage: ConfigSyncService.hasUserOverride ? "person.fill.checkmark" : "doc.badge.gearshape"
                    )
                    .foregroundColor(ConfigSyncService.hasUserOverride ? .blue : .secondary)

                    Text(ConfigSyncService.userConfigURL.path)
                        .font(.system(.caption2, design: .monospaced))
                        .foregroundColor(.secondary)
                        .lineLimit(1)
                        .truncationMode(.middle)
                        .textSelection(.enabled)
                }
                Spacer()
            }

            HStack(spacing: 10) {
                Button("In Finder öffnen") {
                    let dir = ConfigSyncService.userConfigURL.deletingLastPathComponent()
                    // Ordner erstellen falls nicht vorhanden, dann öffnen
                    try? FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
                    NSWorkspace.shared.open(dir)
                }
                .buttonStyle(.bordered)

                if ConfigSyncService.hasUserOverride {
                    Button("Auf Standard zurücksetzen") {
                        errorMessage = nil
                        do {
                            try ConfigSyncService.resetToDefault()
                            appState.addLog(.info, "Brand Config auf Bundle-Standard zurückgesetzt.")
                        } catch {
                            errorMessage = error.localizedDescription
                        }
                    }
                    .buttonStyle(.bordered)
                    .foregroundColor(.red)
                }
            }

            Text("Lege eine eigene Infront_BrandConfig.json im User-Override-Pfad ab, um die Standard-Konfiguration zu überschreiben.")
                .font(.caption)
                .foregroundColor(.secondary)
        }
    }

    private var manifestSection: some View {
        Section("Manifest") {
            VStack(alignment: .leading, spacing: 6) {
                Text("Katalog-Ordner")
                    .font(.subheadline)
                Text(ManifestService.catalogFolderURL.path)
                    .font(.system(.caption, design: .monospaced))
                    .foregroundColor(.secondary)
                    .textSelection(.enabled)
            }

            HStack(spacing: 10) {
                Button("In Finder öffnen") {
                    NSWorkspace.shared.open(ManifestService.catalogFolderURL)
                }
                .buttonStyle(.bordered)

                Button("Manifest neu installieren") {
                    errorMessage = nil
                    do {
                        try ManifestService.installManifest(port: appState.selectedPort)
                        appState.addLog(.info, "Manifest neu installiert (Port \(appState.selectedPort))")
                    } catch {
                        errorMessage = error.localizedDescription
                        appState.addLog(.error, error.localizedDescription)
                    }
                }
                .buttonStyle(.bordered)
            }
        }
    }
}
