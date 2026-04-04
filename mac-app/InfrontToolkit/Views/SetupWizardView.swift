// SetupWizardView.swift
// Erststart-Assistent (6 Schritte).
// Führt den Nutzer durch: Port-Wahl → Server starten → Manifest installieren
// → Office konfigurieren → Autostart → Fertig

import SwiftUI

struct SetupWizardView: View {

    @EnvironmentObject var appState: AppState
    @State private var portInput: String = ""
    @State private var isBusy: Bool = false
    @State private var stepError: String?

    var body: some View {
        VStack(spacing: 0) {
            headerView
            Divider()
            stepContentView
                .frame(maxWidth: .infinity, maxHeight: .infinity)
                .padding(24)
            Divider()
            navigationBarView
                .padding(.horizontal, 20)
                .padding(.vertical, 14)
        }
        .frame(width: 460, height: 560)
        .onAppear { portInput = String(appState.selectedPort) }
    }

    // MARK: - Header

    private var headerView: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Infront Toolkit einrichten")
                .font(.title2).bold()
            ProgressView(
                value: Double(appState.currentSetupStep.rawValue),
                total: Double(SetupStep.allCases.count - 1)
            )
            .tint(.accentColor)
            Text("Schritt \(appState.currentSetupStep.rawValue + 1) von \(SetupStep.allCases.count): \(appState.currentSetupStep.title)")
                .font(.caption)
                .foregroundColor(.secondary)
        }
        .padding(.horizontal, 24)
        .padding(.vertical, 16)
    }

    // MARK: - Step content

    @ViewBuilder
    private var stepContentView: some View {
        switch appState.currentSetupStep {
        case .welcome:       welcomeStep
        case .startServer:   startServerStep
        case .manifest:      manifestStep
        case .officeCatalog: officeCatalogStep
        case .autostart:     autostartStep
        case .done:          doneStep
        }
    }

    // MARK: - Navigation bar

    private var navigationBarView: some View {
        HStack(spacing: 12) {
            if appState.currentSetupStep.rawValue > 0 && appState.currentSetupStep != .done {
                Button("Zurück") {
                    stepError = nil
                    appState.retreatSetupStep()
                }
                .buttonStyle(.plain)
                .foregroundColor(.secondary)
            }

            Spacer()

            if let error = stepError {
                Text(error)
                    .foregroundColor(.red)
                    .font(.caption)
                    .lineLimit(2)
                    .frame(maxWidth: 200, alignment: .trailing)
            }

            primaryActionButton
        }
    }

    @ViewBuilder
    private var primaryActionButton: some View {
        switch appState.currentSetupStep {
        case .done:
            Button("Fertig") {
                appState.completeSetup()
            }
            .buttonStyle(.borderedProminent)

        case .startServer:
            Button(isBusy ? "Startet…" : (appState.serverStatus.isRunning ? "Server läuft ✓" : "Server starten")) {
                Task { await startServerAction() }
            }
            .buttonStyle(.borderedProminent)
            .disabled(isBusy || appState.serverStatus.isRunning)

        case .manifest:
            Button(isBusy ? "Installiert…" : (appState.manifestInstalledPath != nil ? "Erneut installieren" : "Manifest installieren")) {
                installManifestAction()
            }
            .buttonStyle(.borderedProminent)
            .disabled(isBusy)

        default:
            Button("Weiter") {
                stepError = nil
                appState.advanceSetupStep()
            }
            .buttonStyle(.borderedProminent)
        }
    }

    // MARK: - Step 1: Welcome

    private var welcomeStep: some View {
        VStack(alignment: .leading, spacing: 20) {
            HStack {
                Spacer()
                Image(systemName: "bolt.circle.fill")
                    .font(.system(size: 52))
                    .foregroundStyle(.blue)
                Spacer()
            }

            Text("Willkommen beim Infront Toolkit")
                .font(.headline)

            Text("Dieser Assistent richtet das PowerPoint Add-in auf deinem Mac ein in ca. 5 Minuten.\n\nVoraussetzungen:\n• PowerPoint für Mac (Microsoft 365)\n• macOS 13 oder neuer")
                .foregroundColor(.secondary)

            Divider()

            VStack(alignment: .leading, spacing: 6) {
                Text("Port für den lokalen Server:")
                    .font(.subheadline).fontWeight(.medium)
                HStack(spacing: 8) {
                    TextField("3000", text: $portInput)
                        .textFieldStyle(.roundedBorder)
                        .frame(width: 80)
                        .onChange(of: portInput) { val in
                            if let p = Int(val) { appState.savePort(p) }
                        }
                    Text("Standard: 3000")
                        .font(.caption)
                        .foregroundColor(.secondary)
                }
            }
        }
    }

    // MARK: - Step 2: Start Server

    private var startServerStep: some View {
        VStack(alignment: .leading, spacing: 20) {
            HStack {
                Spacer()
                Image(systemName: "server.rack")
                    .font(.system(size: 52))
                    .foregroundStyle(appState.serverStatus.isRunning ? .green : .blue)
                Spacer()
            }

            Text("Lokalen Server starten")
                .font(.headline)

            Text("Das Add-in wird von einem lokalen HTTP-Server auf Port \(appState.selectedPort) bereitgestellt. Der Server muss immer laufen, wenn du das Add-in in PowerPoint verwendest.")
                .foregroundColor(.secondary)

            if appState.serverStatus.isRunning {
                Label("Server läuft auf Port \(appState.selectedPort)", systemImage: "checkmark.circle.fill")
                    .foregroundColor(.green)
                    .fontWeight(.medium)
            } else if case .error(let msg) = appState.serverStatus {
                Label(msg, systemImage: "xmark.circle.fill")
                    .foregroundColor(.red)
                    .font(.caption)
            }
        }
    }

    // MARK: - Step 3: Manifest

    private var manifestStep: some View {
        VStack(alignment: .leading, spacing: 20) {
            HStack {
                Spacer()
                Image(systemName: "doc.badge.plus")
                    .font(.system(size: 52))
                    .foregroundStyle(.blue)
                Spacer()
            }

            Text("Manifest-Datei installieren")
                .font(.headline)

            Text("Das Add-in-Manifest wird in einen lokalen Ordner kopiert. Diesen Ordner registrierst du im nächsten Schritt in Office.")
                .foregroundColor(.secondary)

            if let path = appState.manifestInstalledPath {
                VStack(alignment: .leading, spacing: 4) {
                    Label("Installiert", systemImage: "checkmark.circle.fill")
                        .foregroundColor(.green)
                        .fontWeight(.medium)
                    Text(path.path)
                        .font(.system(.caption, design: .monospaced))
                        .foregroundColor(.secondary)
                        .textSelection(.enabled)
                }
                .padding(10)
                .background(Color.secondary.opacity(0.08))
                .cornerRadius(8)
            }
        }
    }

    // MARK: - Step 4: Office Catalog

    private var officeCatalogStep: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 16) {
                HStack {
                    Spacer()
                    Image(systemName: "doc.text.magnifyingglass")
                        .font(.system(size: 42))
                        .foregroundStyle(.blue)
                    Spacer()
                }

                Text("Office konfigurieren")
                    .font(.headline)

                Text("Registriere den Manifest-Ordner einmalig in Office:")
                    .foregroundColor(.secondary)

                VStack(alignment: .leading, spacing: 10) {
                    ForEach(Array(officeSetupSteps.enumerated()), id: \.offset) { idx, step in
                        HStack(alignment: .top, spacing: 10) {
                            Text("\(idx + 1)")
                                .font(.caption.bold())
                                .foregroundColor(.white)
                                .frame(width: 20, height: 20)
                                .background(Color.accentColor)
                                .clipShape(Circle())
                            Text(step)
                                .font(.subheadline)
                        }
                    }
                }
                .padding(14)
                .background(Color.secondary.opacity(0.08))
                .cornerRadius(10)

                // Pfad-Anzeige + Finder-Shortcut (immer sichtbar)
                VStack(alignment: .leading, spacing: 6) {
                    Text("Katalog-Ordner:")
                        .font(.caption).foregroundColor(.secondary)
                    Text(ManifestService.catalogFolderURL.path)
                        .font(.system(.caption, design: .monospaced))
                        .foregroundColor(.secondary)
                        .textSelection(.enabled)
                        .lineLimit(2)
                }
                .padding(10)
                .background(Color.secondary.opacity(0.06))
                .cornerRadius(8)

                HStack(spacing: 10) {
                    Button("Ordner in Finder öffnen") {
                        NSWorkspace.shared.open(ManifestService.catalogFolderURL)
                    }
                    .buttonStyle(.bordered)

                    Button("Pfad kopieren") {
                        NSPasteboard.general.clearContents()
                        NSPasteboard.general.setString(ManifestService.catalogFolderURL.path, forType: .string)
                    }
                    .buttonStyle(.bordered)
                }
            }
        }
    }

    // MARK: - Step 5: Autostart

    private var autostartStep: some View {
        VStack(alignment: .leading, spacing: 20) {
            HStack {
                Spacer()
                Image(systemName: "clock.arrow.circlepath")
                    .font(.system(size: 52))
                    .foregroundStyle(.blue)
                Spacer()
            }

            Text("Automatisch starten")
                .font(.headline)

            Text("Damit der Server immer läuft wenn PowerPoint geöffnet wird, empfehlen wir den automatischen Start beim macOS-Login.")
                .foregroundColor(.secondary)

            Toggle(
                "Bei Login automatisch starten",
                isOn: Binding(
                    get: { appState.autoStartEnabled },
                    set: { enabled in
                        do {
                            if enabled { try AppAutoStartService.enable() }
                            else       { try AppAutoStartService.disable() }
                            appState.saveAutoStart(enabled)
                        } catch {
                            stepError = error.localizedDescription
                        }
                    }
                )
            )
            .toggleStyle(.switch)
        }
    }

    // MARK: - Step 6: Done

    private var doneStep: some View {
        VStack(spacing: 24) {
            Spacer()
            Image(systemName: "checkmark.seal.fill")
                .font(.system(size: 72))
                .foregroundStyle(.green)
            Text("Einrichtung abgeschlossen!")
                .font(.title2).bold()
            Text("Öffne PowerPoint — der Tab \"Infront Toolkit\" erscheint im Ribbon.\n\nDie App läuft als Menu Bar Icon und startet den Server automatisch.")
                .foregroundColor(.secondary)
                .multilineTextAlignment(.center)
            Spacer()
        }
        .frame(maxWidth: .infinity)
    }

    // MARK: - Actions

    private func startServerAction() async {
        isBusy = true
        defer { isBusy = false }
        await appState.startServer()
        if appState.serverStatus.isRunning {
            try? await Task.sleep(nanoseconds: 500_000_000)
            appState.advanceSetupStep()
        } else if case .error(let msg) = appState.serverStatus {
            stepError = msg
        }
    }

    private func installManifestAction() {
        isBusy = true
        defer { isBusy = false }
        do {
            let url = try ManifestService.installManifest(port: appState.selectedPort)
            appState.manifestInstalledPath = url
            appState.addLog(.info, "Manifest installiert: \(url.path)")
            try? Task.sleep(nanoseconds: 300_000_000) // kurze Pause für UI-Feedback
            appState.advanceSetupStep()
        } catch {
            stepError = error.localizedDescription
            appState.addLog(.error, "Manifest-Installation fehlgeschlagen: \(error.localizedDescription)")
        }
    }

    // MARK: - Helpers

    private var officeSetupSteps: [String] {
        [
            "PowerPoint (oder Word/Excel) öffnen",
            "Menü: Einfügen → Add-ins → Weitere Add-ins…",
            "Tab \"Freigegebene Ordnerkataloge\" wählen",
            "Katalog-URL: den Pfad unten einfügen → \"Katalog hinzufügen\"",
            "\"Zum Menü hinzufügen\" aktivieren → OK",
            "Office neu starten — Tab \"Infront Toolkit\" erscheint im Ribbon"
        ]
    }
}
