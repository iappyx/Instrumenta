# Installing the Instrumenta toolbar

There are **two ways** to install Instrumenta:
- **[Manual installation](#manual-installation)** (recommended) – This method allows you to install the plugin without an installer, requiring no administrative rights on most enterprise systems.
- **[Installer-based installation](#installer-based-installation-beta)** (beta) – This method uses automated installers but may require bypassing security warnings.

---

## Manual installation

Instrumenta is a Visual Basic for Applications (VBA) add-in that can be installed within PowerPoint, requiring no administrative rights on most enterprise systems.

### Windows manual installation

You can save the add-in to your PC and then install it by adding it to the Available Add-Ins list:

- Download the add-in file:  
  - Choose the [latest release](https://github.com/iappyx/Instrumenta/releases/download/1.49/InstrumentaPowerpointToolbar.ppam) or the [latest beta](https://github.com/iappyx/Instrumenta/raw/main/bin/InstrumentaPowerpointToolbar.ppam)  
  - Save it in a fixed location.
- Open PowerPoint, click **File** → **Options**.
- In the **Options** dialog box, click **Add-Ins**.
- In the **Manage** list at the bottom of the dialog box, click **PowerPoint Add-ins**, then click **Go**.
- In the **Add-Ins** dialog box, click **Add New**.
- In the **Add New PowerPoint Add-In** dialog box, browse for the add-in file and click **OK**.
- A security notice appears. Click **Enable Macros**, then **Close**.  
  - *Note:* If you cannot enable Macros in this dialog, follow [Microsoft's instructions](https://support.microsoft.com/en-gb/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) to unblock Instrumenta.

Once installed, there should be an **Instrumenta** page in the PowerPoint ribbon. 
To upgrade to a newer version see: [Upgrading a manual installation](#upgrading-a-manual-installation-windows--macos). 

### macOS manual installation

You can save the add-in to your Mac and install it manually:

- Download the add-in file in the latest release:  
  - Choose the [latest release](https://github.com/iappyx/Instrumenta/releases/download/1.49/InstrumentaPowerpointToolbar.ppam) or the [latest beta](https://github.com/iappyx/Instrumenta/raw/main/bin/InstrumentaPowerpointToolbar.ppam)  
  - Save it in a fixed location.
- Open PowerPoint, click **Tools** → **Add-ins...**.
- In the **Add-Ins** dialog box, click the **+** button, browse for the add-in file, then click **Open**.
- Click **OK** to close the Add-Ins dialog box.
- There should now be an **Instrumenta** page in the PowerPoint ribbon.

#### Optional steps to show group titles:
- By default, group titles are hidden in the ribbon on macOS.
- Open **PowerPoint**, then click **PowerPoint** in the application menu.
- Click **Preferences** → **View**.
- Check **Show group titles**.

#### Optional steps to enable export to Outlook and Word:
- Download the AppleScript file:  
  - [Instrumenta AppleScript Plugin](https://github.com/iappyx/Instrumenta/releases/download/1.49/InstrumentaAppleScriptPlugin.applescript).
- Copy the AppleScript file to: `~/Library/Application Scripts/com.microsoft.Powerpoint/`
- This is in the library folder of the **current user**. If the folder does not exist, create it.
- In some cases, a **reboot** of your Mac might be required.

The AppleScript file will not change often—only when major updates to export features occur. Instrumenta will notify you when an updated version is required.

### Upgrading a manual installation (Windows & MacOS)
To upgrade make sure Powerpoint is closed and overwrite the add-in file with the new version. 

#### Upgrading from version < 1.44
Starting from version 1.44, Instrumenta introduces a tabbed interface within the ribbon, optimizing usability for users without widescreen displays. Features are now organized into multiple tabs based on functionality, including text, shapes, tables, and advanced options.

By default, the new tabbed layout is enabled, providing a more structured experience. However, users who prefer the classic single-tab view can still access it in both Pro and Review modes by adjusting the Instrumenta settings.

If you're upgrading from an earlier version, the classic view will remain active initially. To switch to the new interface, navigate to Instrumenta settings and set the Operating Mode to "Default: Multiple Tabs."

![Alt text](img/instrumenta-win-settings-1.44.png?raw=true "Instrumenta Powerpoint Toolbar settings (Windows)")

#### Upgrading from version < 1.0 to version 1.0 or higher
Please know that all versions before the first major release (v1.0) have a different filename. If you are upgrading to v1.0 please remove the old add-in first in the Add-ins dialog box in Powerpoint and then install the new version as mentioned above.

---

## Installer-based installation (Beta)

Instrumenta offers dedicated installers for both macOS and Windows. You can choose to install Instrumenta Keys alongside the core package. The Windows installer should be able to work without requiring administrative rights, ensuring installation even in restricted environments. The installers are currently in beta and being tested across multiple scenarios to ensure reliability and compatibility. However, as they are still under development, use them at your own risk, and be aware that unexpected issues may arise.

Instrumenta is a PowerPoint add-in that I developed as an open-source project in my free time. Since I do not have official code-signing certificates for Windows or a developer account for macOS, installing through these installers may trigger security warnings.

### Windows installer-based installation

Download and open the [installer for Windows](https://github.com/iappyx/Instrumenta/raw/main/bin/Installers/InstrumentaPowerpointToolbarSetup.exe). 
When running the installer, Windows SmartScreen may block execution. Follow these steps to bypass it:

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

#### Uninstall
You can uninstall Instrumenta through the standard Windows **Add or Remove Programs** menu.

### macOS installer-based installation

Download and open the [Instrumenta Installer Disk Image (dmg)](https://github.com/iappyx/Instrumenta/raw/main/bin/Installers/InstrumentaInstaller.dmg).
 
macOS may prevent the execution of the unsigned `InstrumentaInstaller` installer package within the `InstrumentaInstaller.dmg`. 

![image](https://github.com/user-attachments/assets/b1554c82-db0a-4b24-9252-b5300fbb2557)

#### If the `InstrumentaInstaller` package doesn't open:
1. Open **System Settings** → **Privacy & Security**.
2. Scroll down to the **Security** section.
3. You should see a message stating that `InstrumentaInstaller` was blocked.
4. Click **Allow Anyway**.
5. Try opening `InstrumentaInstaller` again.

If you don't want to allow this, the `InstrumentaInstaller.dmg` also contains a folder called `Files for manual installation`, which has all the files needed for a [manual installation](#manual-installation).

#### Uninstall
You can uninstall Instrumenta by running the installer again. The installer includes a built-in uninstallation option.
