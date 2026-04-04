Attribute VB_Name = "frmMasterImport"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmMasterImport
' Zweck:  Steuert den Infront Master-Importer – Auswahl der Quelldatei,
'         Anzeige verfügbarer Masters, Import mit Optionen.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name               | Typ             | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lblSourceFile      | Label           | Caption = "Quelldatei:"
'  txtSourceFile      | TextBox         | Dateipfad anzeigen (ReadOnly=True)
'  btnBrowse          | CommandButton   | Caption = "Durchsuchen..."
'  btnLoadMasters     | CommandButton   | Caption = "Masters laden"
'  lblMasters         | Label           | Caption = "Verfügbare Masters:"
'  lstMasters         | ListBox         | Zeigt Master-Namen
'  fraOptions         | Frame           | Caption = "Optionen"
'    chkApplyAll      | CheckBox        | Caption = "Master auf alle Folien anwenden"
'    chkApplySelected | CheckBox        | Caption = "Master nur auf selektierte Folien"
'    chkRemoveUnused  | CheckBox        | Caption = "Ungenutzte Masters danach entfernen"
'  lblStatus          | Label           | Statusmeldungen
'  btnImport          | CommandButton   | Caption = "Importieren"
'  btnClose           | CommandButton   | Caption = "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Master-Importer"
'  Width:       340 pt,  Height: 380 pt
'  BorderStyle: 1 (Single)
'
' Hinweis:
'  chkApplyAll und chkApplySelected schließen sich gegenseitig aus –
'  das Aktivieren des einen deaktiviert das andere (siehe Event-Handler).
' =============================================================================

Option Explicit


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form mit Standardwerten.
' -----------------------------------------------------------------------
Public Sub InitForm()

    On Error Resume Next

    txtSourceFile.Text = ""
    lstMasters.Clear
    chkApplyAll.Value = False
    chkApplySelected.Value = False
    chkRemoveUnused.Value = False
    lblStatus.Caption = "Quelldatei auswählen und Masters laden."
    btnImport.Enabled = False

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet Datei-Dialog und befüllt txtSourceFile.
' -----------------------------------------------------------------------
Private Sub btnBrowse_Click()

    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modMasterImport.BrowseForFile()

    If filePath <> "" Then
        txtSourceFile.Text = filePath
        lstMasters.Clear
        btnImport.Enabled = False
        lblStatus.Caption = "Datei ausgewählt. Klicke 'Masters laden'."
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Fehler: " & Err.Description
End Sub


' -----------------------------------------------------------------------
' Zweck:      Lädt Master-Namen aus der Quelldatei und befüllt lstMasters.
' -----------------------------------------------------------------------
Private Sub btnLoadMasters_Click()

    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = Trim(txtSourceFile.Text)

    If filePath = "" Then
        lblStatus.Caption = "Bitte zuerst eine Quelldatei auswählen."
        Exit Sub
    End If

    lblStatus.Caption = "Lade Masters..."
    DoEvents

    Dim masterNames() As String
    Dim masterCount As Long

    If modMasterImport.LoadMastersFromFile(filePath, masterNames, masterCount) Then
        lstMasters.Clear
        Dim i As Long
        For i = 1 To masterCount
            lstMasters.AddItem masterNames(i)
        Next i

        If masterCount > 0 Then
            lstMasters.ListIndex = 0
            btnImport.Enabled = True
            lblStatus.Caption = masterCount & " Master" & _
                IIf(masterCount = 1, "", "s") & " gefunden."
        Else
            btnImport.Enabled = False
            lblStatus.Caption = "Keine Masters in der Datei gefunden."
        End If
    Else
        lblStatus.Caption = "Masters konnten nicht geladen werden."
        btnImport.Enabled = False
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Fehler: " & Err.Description
    btnImport.Enabled = False
End Sub


' -----------------------------------------------------------------------
' Zweck:      chkApplyAll und chkApplySelected schließen sich gegenseitig aus.
' -----------------------------------------------------------------------
Private Sub chkApplyAll_Click()
    If chkApplyAll.Value = True Then
        chkApplySelected.Value = False
    End If
End Sub

Private Sub chkApplySelected_Click()
    If chkApplySelected.Value = True Then
        chkApplyAll.Value = False
    End If
End Sub


' -----------------------------------------------------------------------
' Zweck:      Importiert den gewählten Master.
' -----------------------------------------------------------------------
Private Sub btnImport_Click()

    On Error GoTo ErrHandler

    If lstMasters.ListIndex < 0 Then
        lblStatus.Caption = "Bitte einen Master aus der Liste auswählen."
        Exit Sub
    End If

    Dim filePath As String
    filePath = Trim(txtSourceFile.Text)

    If filePath = "" Then
        lblStatus.Caption = "Keine Quelldatei angegeben."
        Exit Sub
    End If

    ' Master-Index (1-basiert)
    Dim masterIndex As Long
    masterIndex = lstMasters.ListIndex + 1

    ' Optionen
    Dim opts As modMasterImport.ImportOptions
    opts.ApplyToAllSlides       = (chkApplyAll.Value = True)
    opts.ApplyToSelectedSlides  = (chkApplySelected.Value = True)
    opts.RemoveUnusedAfterImport = (chkRemoveUnused.Value = True)

    lblStatus.Caption = "Importiere '" & lstMasters.List(lstMasters.ListIndex) & "'..."
    DoEvents

    If modMasterImport.ImportMaster(filePath, masterIndex, opts) Then
        Dim msg As String
        msg = "Master '" & lstMasters.List(lstMasters.ListIndex) & "' importiert."
        If opts.ApplyToAllSlides Then
            msg = msg & vbCrLf & "Auf alle Folien angewendet."
        ElseIf opts.ApplyToSelectedSlides Then
            msg = msg & vbCrLf & "Auf selektierte Folien angewendet."
        End If
        lblStatus.Caption = msg
    Else
        lblStatus.Caption = "Import fehlgeschlagen – Details siehe Meldung."
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Fehler: " & Err.Description
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form.
' -----------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form sauber bei X-Button / Alt+F4.
' -----------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = False
    End If
End Sub
