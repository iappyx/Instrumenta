// StatusMenuView.swift
// Menu Bar Dropdown – erscheint beim Klick auf das Bolt-Icon in der Menu Bar.
// Nur Button + Divider erlaubt (MenuBarExtra mit .menu Style).

import SwiftUI

struct StatusMenuView: View {

    @EnvironmentObject var appState: AppState
    @Environment(\.openWindow) private var openWindow

    var body: some View {
        // Statuszeile (deaktivierter Button als Label-Ersatz)
        Button(appState.serverStatus.description) {}
            .disabled(true)

        Divider()

        Button("Dashboard öffnen…") {
            openWindow(id: "main-window")
            NSApplication.shared.activate(ignoringOtherApps: true)
        }

        Divider()

        if appState.serverStatus.isRunning {
            Button("Server stoppen") {
                Task { await appState.stopServer() }
            }
        } else {
            Button("Server starten") {
                Task { await appState.startServer() }
            }
        }

        Divider()

        Button("Beenden") {
            NSApplication.shared.terminate(nil)
        }
    }
}
