!include MUI2.nsh

InstType "Typical"
InstType "Complete" 

!define MUI_COMPONENTSPAGE
!define MUI_ICON "..\..\..\..\Instrumenta-Keys\src\windows\Instrumenta.ico"
!define MUI_UNICON "..\..\..\..\Instrumenta-Keys\src\windows\Instrumenta.ico"
!define MUI_COMPONENTSPAGE_TEXT "Select optional components to install."

!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES 
!include LogicLib.nsh

!insertmacro MUI_LANGUAGE "English"

Var HOST_ARCH
Var INSTRUMENTA_RELEASE
Var CURRENT_OFFICE_RELEASE
Var CURRENT_OFFICE_REGKEY
Var ADDIN_FILE
Var CONFIGURED
Var INSTALL_DIR

Name "Instrumenta Powerpoint Toolbar"
OutFile "..\..\..\bin\Installers\InstrumentaPowerpointToolbarSetup.exe"

RequestExecutionLevel user

Function .onInit
    StrCpy $INSTRUMENTA_RELEASE "1.0"
    StrCpy $CONFIGURED ""
    ReadRegStr $HOST_ARCH HKLM "System\CurrentControlSet\Control\Session Manager\Environment" "PROCESSOR_ARCHITECTURE"
    ${If} $HOST_ARCH == "AMD64"
        SetRegView 64  
    ${Else}
        SetRegView 32  
    ${EndIf}
FunctionEnd

Function Handle_Addin_Registration
    Push $0
    StrCpy $0 0

Loop_Office_Version:
    EnumRegKey $CURRENT_OFFICE_RELEASE HKLM "Software\Microsoft\Office" $0
    StrCmp $CURRENT_OFFICE_RELEASE "" NoOfficeFound

    StrCpy $CURRENT_OFFICE_REGKEY "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\InstallRoot"
    ReadRegStr $INSTALL_DIR HKLM "$CURRENT_OFFICE_REGKEY" "Path"
    StrCmp $INSTALL_DIR "" Next_Office 0 

    StrCmp $CURRENT_OFFICE_RELEASE "16.0" OfficeDetected
    Goto Next_Office  

OfficeDetected:
    StrCpy $CONFIGURED "yes"
    StrCpy $ADDIN_FILE "$APPDATA\Microsoft\AddIns\InstrumentaPowerpointToolbar.ppam"
    Call RegisterAddin

    Return  

Next_Office:
    IntOp $0 $0 + 1
    Goto Loop_Office_Version

NoOfficeFound:
    MessageBox MB_OK "No compatible Office version found! Office 2016+ required. You might try a manual install. See: https://github.com/iappyx/Instrumenta/blob/main/README.md"
    Quit
FunctionEnd

Function RegisterAddin
    WriteRegStr HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\InstrumentaPowerpointToolbar" "Path" "$ADDIN_FILE"
    WriteRegDWORD HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\InstrumentaPowerpointToolbar" "AutoLoad" 1
    WriteRegStr HKCU "Software\Instrumenta\OfficeVersion" "Version" "$CURRENT_OFFICE_RELEASE"   
    DetailPrint "Add-in registered successfully!"
FunctionEnd

Section "Instrumenta" MainSection

    SectionIn RO
    Call Handle_Addin_Registration
    SetOutPath "$APPDATA\Microsoft\AddIns\"
    File "..\..\..\bin\InstrumentaPowerpointToolbar.ppam"

    SetOutPath "$LOCALAPPDATA\Instrumenta\"
    File "..\..\..\..\Instrumenta-Keys\src\windows\Instrumenta.ico"
    WriteUninstaller "$LOCALAPPDATA\Instrumenta\InstrumentaPowerpointToolbar-uninstall.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "DisplayIcon" "$LOCALAPPDATA\Instrumenta\Instrumenta.ico"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "DisplayName" "Instrumenta PowerPoint Toolbar"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "UninstallString" '"$LOCALAPPDATA\Instrumenta\InstrumentaPowerpointToolbar-uninstall.exe"'
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "InstallLocation" "$LOCALAPPDATA\Instrumenta"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "Publisher" "Iappyx"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "DisplayVersion" "1.50" 
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "HelpLink" "https://github.com/iappyx/Instrumenta/blob/main/README.md"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "URLInfoAbout" "https://github.com/iappyx/Instrumenta"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "Comments" "Instrumenta is a free and open source consulting-style PowerPoint toolbar"
    WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "EstimatedSize" 1024

    WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "NoModify" 1
    WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar" "NoRepair" 1

    DetailPrint "Installation complete! Restart PowerPoint to activate the add-in."
    
SectionEnd

Section "Keys (optional)" OptionalSection
    SectionSetText ${OptionalSection} "Installs the Instrumenta PowerPoint Toolbar for enhanced functionality." 
    SectionIn 2
    SetOutPath "$LOCALAPPDATA\Instrumenta\"
    File "..\..\..\..\Instrumenta-Keys\bin\windows\Instrumenta Keys.exe"
    CreateShortcut "$DESKTOP\Instrumenta Keys.lnk" "$LOCALAPPDATA\Instrumenta\Instrumenta Keys.exe"
SectionEnd

LangString DESC_MainSection ${LANG_ENGLISH} "Instrumenta is a free and open source consulting-style PowerPoint toolbar."
LangString DESC_OptionalSection ${LANG_ENGLISH} "Instrumenta Keys is a keyboard shortcut companion for Instrumenta, bringing customizable keyboard shortcuts to Instrumenta."

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${MainSection} $(DESC_MainSection)
    !insertmacro MUI_DESCRIPTION_TEXT ${OptionalSection} $(DESC_OptionalSection)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Section "Uninstall"

    ReadRegStr $HOST_ARCH HKLM "System\CurrentControlSet\Control\Session Manager\Environment" "PROCESSOR_ARCHITECTURE"
    ${If} $HOST_ARCH == "AMD64"
    SetRegView 64  
    ${Else}
    SetRegView 32  
    ${EndIf}

    MessageBox MB_YESNO "Do you want to uninstall Instrumenta? Removing it will delete all configurations and keyboard shortcuts." IDNO cancel_uninstall

    Delete "$APPDATA\Microsoft\AddIns\InstrumentaPowerpointToolbar.ppam"
    Delete "$LOCALAPPDATA\Instrumenta\Instrumenta Keys.exe"
    Delete "$LOCALAPPDATA\Instrumenta\shortcuts.csv"
    Delete "$LOCALAPPDATA\Instrumenta\InstrumentaPowerpointToolbar-uninstall.exe"
    RMDir /r "$LOCALAPPDATA\Instrumenta\"
    Delete "$DESKTOP\Instrumenta Keys.lnk"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\InstrumentaPowerpointToolbar"

    ReadRegStr $CURRENT_OFFICE_RELEASE HKCU "Software\Instrumenta\OfficeVersion" "Version"
    DeleteRegKey HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\InstrumentaPowerpointToolbar"
    DeleteRegKey HKCU "Software\Instrumenta\OfficeVersion"
    DeleteRegKey HKCU "Software\Instrumenta"

    MessageBox MB_OK "InstrumentaPowerpointToolbar has been removed!"
    Quit

    cancel_uninstall:
    MessageBox MB_OK "Uninstallation canceled."
    Quit
SectionEnd
