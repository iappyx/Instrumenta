# Installer-based installation of Instrumenta

Instrumenta is a PowerPoint plugin that I developed as an open-source project in my free time. Since I do not have official code-signing certificates for Windows or a developer account for macOS, installing the plugin may trigger security warnings. This guide provides step-by-step instructions to bypass these warnings and install the software through the installers.

You can always choose not to use the installers and do a manual installation instead.

## Windows installation

When running the `InstrumentaPowerpointToolbarSetup.exe` installer, Windows SmartScreen may block execution. Follow these steps to bypass it:

1. **Run the installer**  
   Open `InstrumentaPowerpointToolbarSetup.exe`.

2. **SmartScreen warning appears**  
   If a security prompt appears stating *"Windows protected your PC"*, do the following:

   - Click **More info**.
   - You will see the publisher as "Unknown" since the installer is unsigned.
   - Click the **Run anyway** button.

3. **Confirm Installation**  
   Follow the installation steps as guided.

![image](https://github.com/user-attachments/assets/43bba1eb-6e30-4fd2-a16e-eb162fe62b1a)

![image](https://github.com/user-attachments/assets/7456d161-d17b-4286-b656-57f0d9cf1d37)


---

## MacOS installation

macOS may prevent the execution of the unsigned `InstrumentaInstaller` installer package within the `InstrumentaInstaller.dmg`. 

![image](https://github.com/user-attachments/assets/b1554c82-db0a-4b24-9252-b5300fbb2557)

### If the `InstrumentaInstaller` package doesn't open
1. Open **System Settings** → **Privacy & Security**.
2. Scroll down to the **Security** section.
3. You should see a message stating that `InstrumentaInstaller` was blocked.
4. Click **Allow Anyway**.
5. Try opening `InstrumentaInstaller` again.

If you don't want to allow this, the `InstrumentaInstaller.dmg` also contains the folder `Files for manual installation`. There you'll find all the files you'll need for a `manual installation`.
