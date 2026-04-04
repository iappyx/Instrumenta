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
