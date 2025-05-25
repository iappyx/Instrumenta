# Installer-based installation of Instrumenta

Instrumenta is a PowerPoint plugin that I developed as an open-source project in my free time. Since I do not have official code-signing certificates for Windows or a developer account for macOS, installing the plugin may trigger security warnings. This guide provides step-by-step instructions to bypass these warnings and install the software trough the installers.

You can always choose not to use the installers and do a manual installation instead.

## Windows

When running the `InstrumentaPowerpointToolbarSetup.exe` installer, Windows SmartScreen may block execution. Follow these steps to bypass it:

1. **Run the installer**  
   Double-click `InstrumentaPowerpointToolbarSetup.exe`.

2. **SmartScreen warning appears**  
   If a security prompt appears stating *"Windows protected your PC"*, do the following:

   - Click **More info**.
   - You will see the publisher as "Unknown" since the installer is unsigned.
   - Click the **Run anyway** button.

3. **Confirm Installation**  
   Follow the installation steps as guided.

---

## Mac

macOS may prevent the execution of the unsigned `.pkg` installer within the DMG. Here’s how to install `InstrumentaInstaller.pkg`:

### If the `.pkg` Installer doesn't open
1. Open **System Settings** → **Privacy & Security**.
2. Scroll down to the **Security** section.
3. You should see a message stating that `InstrumentaInstaller.pkg` was blocked.
4. Click **Allow Anyway**.
5. Try opening `InstrumentaInstaller.pkg` again.

