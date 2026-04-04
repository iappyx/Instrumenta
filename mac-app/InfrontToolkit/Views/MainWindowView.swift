// MainWindowView.swift
// Hauptfenster der App. Zeigt Setup-Assistent (Erststart) oder Dashboard (danach).

import SwiftUI

struct MainWindowView: View {

    @EnvironmentObject var appState: AppState

    var body: some View {
        if !appState.setupCompleted {
            SetupWizardView()
                .environmentObject(appState)
        } else {
            DashboardView()
                .environmentObject(appState)
        }
    }
}

// MARK: - Dashboard (nach Setup)

struct DashboardView: View {

    @EnvironmentObject var appState: AppState

    var body: some View {
        TabView {
            StatusTabView()
                .tabItem { Label("Status", systemImage: "bolt.circle") }
            SettingsView()
                .tabItem { Label("Einstellungen", systemImage: "gear") }
            LogsView()
                .tabItem { Label("Logs", systemImage: "doc.text") }
        }
        .environmentObject(appState)
        .frame(minWidth: 460, minHeight: 480)
    }
}

// MARK: - Status Tab

struct StatusTabView: View {

    @EnvironmentObject var appState: AppState

    var body: some View {
        ScrollView {
            VStack(spacing: 16) {
                serverCard
                manifestCard
                Spacer()
            }
            .padding(20)
        }
    }

    private var serverCard: some View {
        VStack(alignment: .leading, spacing: 12) {
            HStack(spacing: 8) {
                Circle()
                    .fill(appState.serverStatus.statusColor)
                    .frame(width: 10, height: 10)
                Text(appState.serverStatus.description)
                    .font(.headline)
                Spacer()
            }

            HStack(spacing: 10) {
                if appState.serverStatus.isRunning {
                    Button("Stoppen") {
                        Task { await appState.stopServer() }
                    }
                    .buttonStyle(.bordered)

                    Link(
                        "http://localhost:\(appState.selectedPort)",
                        destination: URL(string: "http://localhost:\(appState.selectedPort)")!
                    )
                    .font(.caption)
                } else {
                    Button("Server starten") {
                        Task { await appState.startServer() }
                    }
                    .buttonStyle(.borderedProminent)
                }
                Spacer()
            }
        }
        .padding(16)
        .background(Color.secondary.opacity(0.07))
        .cornerRadius(10)
    }

    private var manifestCard: some View {
        VStack(alignment: .leading, spacing: 10) {
            Text("Manifest")
                .font(.subheadline).bold()

            if ManifestService.isInstalled {
                Label("Installiert", systemImage: "checkmark.circle.fill")
                    .foregroundColor(.green)
                Text(ManifestService.installedManifestURL.path)
                    .font(.system(.caption, design: .monospaced))
                    .foregroundColor(.secondary)
                    .textSelection(.enabled)
                    .lineLimit(2)
            } else {
                Label("Nicht installiert", systemImage: "xmark.circle")
                    .foregroundColor(.orange)
                Button("Manifest installieren") {
                    do {
                        let url = try ManifestService.installManifest(port: appState.selectedPort)
                        appState.addLog(.info, "Manifest installiert: \(url.path)")
                    } catch {
                        appState.addLog(.error, error.localizedDescription)
                    }
                }
                .buttonStyle(.bordered)
            }
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(16)
        .background(Color.secondary.opacity(0.07))
        .cornerRadius(10)
    }
}

// MARK: - Logs View

struct LogsView: View {

    @EnvironmentObject var appState: AppState

    var body: some View {
        VStack(alignment: .leading, spacing: 0) {
            HStack {
                Text("\(appState.logEntries.count) Einträge")
                    .font(.caption)
                    .foregroundColor(.secondary)
                Spacer()
                Button("Leeren") { appState.logEntries.removeAll() }
                    .font(.caption)
                    .buttonStyle(.plain)
                    .foregroundColor(.secondary)
            }
            .padding(.horizontal, 12)
            .padding(.vertical, 8)

            Divider()

            ScrollViewReader { proxy in
                ScrollView {
                    LazyVStack(alignment: .leading, spacing: 2) {
                        ForEach(appState.logEntries) { entry in
                            logRow(entry).id(entry.id)
                        }
                    }
                    .padding(10)
                }
                .onChange(of: appState.logEntries.count) { _ in
                    if let last = appState.logEntries.last {
                        proxy.scrollTo(last.id, anchor: .bottom)
                    }
                }
            }
        }
    }

    private func logRow(_ entry: LogEntry) -> some View {
        HStack(alignment: .top, spacing: 8) {
            Text(entry.formattedTimestamp)
                .font(.system(.caption, design: .monospaced))
                .foregroundColor(.secondary)
                .frame(width: 56, alignment: .leading)
            Text(entry.level.rawValue)
                .font(.system(.caption, design: .monospaced))
                .foregroundColor(entry.level.color)
                .frame(width: 32, alignment: .leading)
            Text(entry.message)
                .font(.system(.caption, design: .monospaced))
        }
    }
}
