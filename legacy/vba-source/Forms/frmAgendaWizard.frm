Attribute VB_Name = "frmAgendaWizard"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmAgendaWizard
' Zweck:  Steuert den Agenda Wizard – Konfiguration und Generierung von
'         Übersichts- und optionalen Fortschrittsfolien.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name               | Typ             | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lblTitle           | Label           | Caption = "Folientitel:"
'  txtTitle           | TextBox         | Agendatitel (z.B. "Agenda")
'  lblItems           | Label           | Caption = "Agendapunkte (1 pro Zeile):"
'  txtItems           | TextBox         | MultiLine=True, ScrollBars=2 (Vertical)
'  fraColors          | Frame           | Caption = "Farben"
'    lblActiveColor   | Label           | Caption = "Aktiv:"
'    txtActiveColor   | TextBox         | Hex-Wert, z.B. #003366
'    btnPickActive    | CommandButton   | Caption = "..."
'    lblInactiveColor | Label           | Caption = "Inaktiv:"
'    txtInactiveColor | TextBox         | Hex-Wert, z.B. #B4B4B4
'    btnPickInactive  | CommandButton   | Caption = "..."
'    lblDoneColor     | Label           | Caption = "Erledigt:"
'    txtDoneColor     | TextBox         | Hex-Wert, z.B. #646464
'    btnPickDone      | CommandButton   | Caption = "..."
'  fraFontSizes       | Frame           | Caption = "Schriftgrößen (pt)"
'    lblTitleSize     | Label           | Caption = "Titel:"
'    txtTitleSize     | TextBox         | Zahl, z.B. 24
'    lblItemSize      | Label           | Caption = "Punkte:"
'    txtItemSize      | TextBox         | Zahl, z.B. 16
'  fraMode            | Frame           | Caption = "Einfügemodus"
'    optOverviewOnly  | OptionButton    | Caption = "Nur Übersichtsfolie"
'    optWithProgress  | OptionButton    | Caption = "Übersicht + Fortschrittsfolien"
'  lblInsertAfter     | Label           | Caption = "Einfügen nach Folie Nr.:"
'  txtInsertAfter     | TextBox         | Zahl (0 = Anfang)
'  lblStatus          | Label           | Statusmeldungen
'  btnGenerate        | CommandButton   | Caption = "Generieren"
'  btnDelete          | CommandButton   | Caption = "Agenda löschen"
'  btnClose           | CommandButton   | Caption = "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Agenda Wizard"
'  Width:       340 pt,  Height: 500 pt
'  BorderStyle: 1 (Single)
' =============================================================================

Option Explicit

' Standard-Farben
Private Const DEFAULT_ACTIVE   As String = "#003366"
Private Const DEFAULT_INACTIVE As String = "#B4B4B4"
Private Const DEFAULT_DONE     As String = "#646464"


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form mit Standardwerten.
'             Wird von ShowAgendaWizard aufgerufen.
' -----------------------------------------------------------------------
Public Sub InitForm()

    On Error Resume Next

    txtTitle.Text = "Agenda"
    txtItems.Text = ""

    txtActiveColor.Text   = DEFAULT_ACTIVE
    txtInactiveColor.Text = DEFAULT_INACTIVE
    txtDoneColor.Text     = DEFAULT_DONE

    txtTitleSize.Text = "24"
    txtItemSize.Text  = "16"
    txtInsertAfter.Text = "0"

    optOverviewOnly.Value = True

    ' Vorhandene Agenda-Folien anzeigen
    UpdateStatus

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Liest Form aus, baut AgendaConfig, ruft GenerateAgenda auf.
' -----------------------------------------------------------------------
Private Sub btnGenerate_Click()

    On Error GoTo ErrHandler

    ' Items parsen
    Dim cfg As modAgendaWizard.AgendaConfig
    modAgendaWizard.ParseItemList txtItems.Text, cfg.Items, cfg.ItemCount

    If cfg.ItemCount = 0 Then
        lblStatus.Caption = "Bitte mindestens einen Agendapunkt eingeben."
        Exit Sub
    End If

    cfg.Title = Trim(txtTitle.Text)

    ' Farben aus Hex-Eingabe lesen
    cfg.ActiveColor   = HexToRGB(txtActiveColor.Text)
    cfg.InactiveColor = HexToRGB(txtInactiveColor.Text)
    cfg.DoneColor     = HexToRGB(txtDoneColor.Text)

    ' Schriftgrößen
    On Error Resume Next
    cfg.TitleFontSize = CSng(txtTitleSize.Text)
    cfg.ItemFontSize  = CSng(txtItemSize.Text)
    On Error GoTo ErrHandler

    ' Einfügemodus
    cfg.InsertionMode = IIf(optWithProgress.Value, 1, 0)

    ' Einfügeposition
    On Error Resume Next
    cfg.InsertAfterSlide = CLng(txtInsertAfter.Text)
    On Error GoTo ErrHandler
    If cfg.InsertAfterSlide < 0 Then cfg.InsertAfterSlide = 0

    lblStatus.Caption = "Wird generiert..."
    DoEvents

    modAgendaWizard.GenerateAgenda cfg

    UpdateStatus

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Fehler: " & Err.Description
End Sub


' -----------------------------------------------------------------------
' Zweck:      Löscht alle Agenda-Folien nach Bestätigung.
' -----------------------------------------------------------------------
Private Sub btnDelete_Click()

    On Error GoTo ErrHandler

    Dim count As Long
    count = modAgendaWizard.CountAgendaSlides()

    If count = 0 Then
        lblStatus.Caption = "Keine Agenda-Folien vorhanden."
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox(count & " Agenda-Folie" & IIf(count = 1, "", "n") & " löschen?", _
                    vbQuestion + vbOKCancel, "Infront Toolkit – Agenda Wizard")

    If answer = vbOK Then
        Dim deleted As Long
        deleted = modAgendaWizard.DeleteExistingAgendaSlides()
        lblStatus.Caption = deleted & " Folie" & IIf(deleted = 1, "", "n") & " gelöscht."
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Fehler: " & Err.Description
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet Farbauswahl für aktive Farbe.
' -----------------------------------------------------------------------
Private Sub btnPickActive_Click()
    Dim picked As String
    picked = PickColorHex(txtActiveColor.Text)
    If picked <> "" Then txtActiveColor.Text = picked
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet Farbauswahl für inaktive Farbe.
' -----------------------------------------------------------------------
Private Sub btnPickInactive_Click()
    Dim picked As String
    picked = PickColorHex(txtInactiveColor.Text)
    If picked <> "" Then txtInactiveColor.Text = picked
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet Farbauswahl für erledigte Farbe.
' -----------------------------------------------------------------------
Private Sub btnPickDone_Click()
    Dim picked As String
    picked = PickColorHex(txtDoneColor.Text)
    If picked <> "" Then txtDoneColor.Text = picked
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


' =============================================================================
' HILFS-FUNKTIONEN (privat)
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Aktualisiert lblStatus mit Anzahl vorhandener Agenda-Folien.
' -----------------------------------------------------------------------
Private Sub UpdateStatus()
    On Error Resume Next
    Dim count As Long
    count = modAgendaWizard.CountAgendaSlides()
    If count = 0 Then
        lblStatus.Caption = "Keine Agenda-Folien vorhanden."
    Else
        lblStatus.Caption = count & " Agenda-Folie" & IIf(count = 1, "", "n") & " vorhanden."
    End If
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet einfachen Farbauswahl-Dialog (InputBox mit Hex-Vorschlag).
'             Auf Windows wird versucht, den nativen Color-Dialog zu nutzen.
' Parameter:  currentHex - aktueller Hex-Wert als Vorschlag
' Rückgabe:   Neuer Hex-Wert "#RRGGBB" oder "" bei Abbruch
' -----------------------------------------------------------------------
Private Function PickColorHex(currentHex As String) As String

    Dim result As String

#If Not Mac Then
    ' Windows: nativer Farbdialog über modColorPicker falls verfügbar
    On Error Resume Next
    Dim currentRGB As Long
    currentRGB = HexToRGB(currentHex)

    ' Versuche Windows-Color-Dialog zu öffnen
    Dim chosenColor As Long
    chosenColor = -1

    ' Fallback auf InputBox wenn Color-Dialog nicht verfügbar
    If chosenColor = -1 Then
        Dim inp As String
        inp = InputBox("Farbe als Hex eingeben (#RRGGBB):", _
                       "Farbe wählen", currentHex)
        If inp <> "" Then result = NormalizeHex(inp)
    End If
    On Error GoTo 0
#Else
    ' Mac: InputBox
    Dim inp As String
    inp = InputBox("Farbe als Hex eingeben (#RRGGBB):", _
                   "Farbe wählen", currentHex)
    If inp <> "" Then result = NormalizeHex(inp)
#End If

    PickColorHex = result
End Function


' -----------------------------------------------------------------------
' Zweck:      Konvertiert #RRGGBB-Hex-String in RGB-Long.
' Parameter:  hex - String wie "#003366" oder "003366"
' Rückgabe:   RGB-Long (0 bei Fehler)
' -----------------------------------------------------------------------
Private Function HexToRGB(hexStr As String) As Long

    On Error GoTo ReturnZero

    Dim clean As String
    clean = Replace(hexStr, "#", "")
    clean = Trim(clean)

    If Len(clean) <> 6 Then GoTo ReturnZero

    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(clean, 2))
    g = CLng("&H" & Mid(clean, 3, 2))
    b = CLng("&H" & Right(clean, 2))

    HexToRGB = RGB(r, g, b)
    Exit Function

ReturnZero:
    HexToRGB = 0
End Function


' -----------------------------------------------------------------------
' Zweck:      Normalisiert Hex-Eingabe auf "#RRGGBB" Format.
' -----------------------------------------------------------------------
Private Function NormalizeHex(inp As String) As String

    Dim clean As String
    clean = Trim(inp)
    If Left(clean, 1) <> "#" Then clean = "#" & clean
    If Len(clean) = 7 Then
        NormalizeHex = UCase(clean)
    Else
        NormalizeHex = inp  ' Ungültig – unverändert zurückgeben
    End If
End Function
