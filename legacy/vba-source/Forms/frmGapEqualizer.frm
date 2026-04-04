Attribute VB_Name = "frmGapEqualizer"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmGapEqualizer
' Zweck:  Steuert den Smart Gap Equalizer – Richtung, Gap-Modus und Anker.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name                | Typ           | Caption / Verwendung
' ---------------------------------------------------------------------------
'  fraDirection        | Frame         | Caption = "Richtung"
'    chkHorizontal     | CheckBox      | Caption = "Horizontal"
'    chkVertical       | CheckBox      | Caption = "Vertikal"
'  fraGapMode          | Frame         | Caption = "Ziel-Abstand"
'    optCustom         | OptionButton  | Caption = "Benutzerdefiniert (pt):"
'    txtGapPt          | TextBox       | Zahl in pt
'    optAverage        | OptionButton  | Caption = "Aktuellen Durchschnitt"
'    optMinimum        | OptionButton  | Caption = "Aktuelles Minimum"
'    optMaximum        | OptionButton  | Caption = "Aktuelles Maximum"
'  fraAnchor           | Frame         | Caption = "Anker"
'    optAnchorFirst    | OptionButton  | Caption = "Erstes Shape fixiert"
'    optAnchorBounds   | OptionButton  | Caption = "Innerhalb Bounds verteilen"
'  lblCurrentInfo      | Label         | Zeigt aktuelle Gap-Werte als Vorschau
'  btnRefresh          | CommandButton | Caption = "Aktualisieren"
'  btnApply            | CommandButton | Caption = "Anwenden"
'  btnClose            | CommandButton | Caption = "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Gap Equalizer"
'  Width:       280 pt,  Height: 340 pt
'  BorderStyle: 1 (Single)
' =============================================================================

Option Explicit


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form.
' -----------------------------------------------------------------------
Public Sub InitForm()

    On Error Resume Next

    chkHorizontal.Value = True
    chkVertical.Value   = False

    optAverage.Value = True
    txtGapPt.Text    = "8"
    txtGapPt.Enabled = False

    optAnchorFirst.Value = True

    RefreshInfo

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Aktualisiert lblCurrentInfo mit aktuellen Gap-Werten.
' -----------------------------------------------------------------------
Private Sub RefreshInfo()

    On Error Resume Next

    Dim info As String

    If chkHorizontal.Value Then
        info = modGapEqualizer.GetGapInfo(True)
    End If

    If chkVertical.Value Then
        Dim vInfo As String
        vInfo = modGapEqualizer.GetGapInfo(False)
        If info <> "" Then
            info = info & vbCrLf & vInfo
        Else
            info = vInfo
        End If
    End If

    If info = "" Then
        info = "Shapes selektieren und aktualisieren."
    End If

    lblCurrentInfo.Caption = info

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      txtGapPt nur aktivieren wenn optCustom UND optAnchorFirst gewählt.
'             Im Bounds-Modus wird der Custom-Wert ignoriert (Gap wird aus
'             Bounds/Shape-Summe berechnet).
' -----------------------------------------------------------------------
Private Sub optCustom_Click()
    txtGapPt.Enabled = (optAnchorFirst.Value = True)
    If txtGapPt.Enabled Then txtGapPt.SetFocus
End Sub

Private Sub optAverage_Click()
    txtGapPt.Enabled = False
End Sub

Private Sub optMinimum_Click()
    txtGapPt.Enabled = False
End Sub

Private Sub optMaximum_Click()
    txtGapPt.Enabled = False
End Sub

Private Sub optAnchorFirst_Click()
    ' Im First-Anchor-Modus ist Custom-Gap nutzbar
    If optCustom.Value Then txtGapPt.Enabled = True
End Sub

Private Sub optAnchorBounds_Click()
    ' Im Bounds-Modus wird der Gap aus Bounds/Shape-Summe berechnet – Custom ignoriert
    txtGapPt.Enabled = False
End Sub


' -----------------------------------------------------------------------
' Zweck:      Aktualisiert Gap-Preview.
' -----------------------------------------------------------------------
Private Sub btnRefresh_Click()
    RefreshInfo
End Sub


' -----------------------------------------------------------------------
' Zweck:      Liest Form aus, baut GapOptions, ruft EqualizeGaps auf.
' -----------------------------------------------------------------------
Private Sub btnApply_Click()

    On Error GoTo ErrHandler

    If Not chkHorizontal.Value And Not chkVertical.Value Then
        lblCurrentInfo.Caption = "Bitte mindestens eine Richtung wählen."
        Exit Sub
    End If

    Dim opts As modGapEqualizer.GapOptions

    opts.Horizontal = (chkHorizontal.Value = True)
    opts.Vertical   = (chkVertical.Value = True)

    ' GapMode
    If optCustom.Value Then
        opts.GapMode = 0
        On Error Resume Next
        opts.CustomGapPt = CSng(txtGapPt.Text)
        If Err.Number <> 0 Then
            lblCurrentInfo.Caption = "Ungültiger Wert in pt-Feld."
            Exit Sub
        End If
        On Error GoTo ErrHandler
    ElseIf optAverage.Value Then
        opts.GapMode = 1
    ElseIf optMinimum.Value Then
        opts.GapMode = 2
    ElseIf optMaximum.Value Then
        opts.GapMode = 3
    End If

    ' AnchorMode
    opts.AnchorMode = IIf(optAnchorFirst.Value, 0, 1)

    modGapEqualizer.EqualizeGaps opts

    ' Info nach Apply aktualisieren
    RefreshInfo

    Exit Sub
ErrHandler:
    lblCurrentInfo.Caption = "Fehler: " & Err.Description
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
