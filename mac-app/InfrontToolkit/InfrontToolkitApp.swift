// InfrontToolkitApp.swift
// Infront Toolkit – native macOS Begleit-App
// Entry Point: Menu Bar App (kein Dock-Icon, gesteuert über LSUIElement in Info.plist)

import SwiftUI

@main
struct InfrontToolkitApp: App {

    @StateObject private var appState = AppState()

    var body: some Scene {
        // Menu Bar Icon + Dropdown
        MenuBarExtra {
            StatusMenuView()
                .environmentObject(appState)
        } label: {
            Image(systemName: appState.serverStatus.menuBarIconName)
                .symbolRenderingMode(.hierarchical)
                .accessibilityLabel("Infront Toolkit")
        }
        .menuBarExtraStyle(.menu)

        // Hauptfenster (Dashboard) – geöffnet per "Dashboard öffnen" im Menü
        Window("Infront Toolkit", id: "main-window") {
            MainWindowView()
                .environmentObject(appState)
        }
        .defaultSize(width: 460, height: 560)
        .windowResizability(.contentSize)
    }
}
