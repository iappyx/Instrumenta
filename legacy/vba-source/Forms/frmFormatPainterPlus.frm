Attribute VB_Name = "frmFormatPainterPlus"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmFormatPainterPlus
' Zweck:  Steuert den Format Painter Plus – wählt welche Eigenschaften vom
'         gecapturten Quell-Shape auf die selektierten Ziel-Shapes übertragen
'         werden.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name              | Typ             | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lblSourceInfo     | Label           | Zeigt Quell-Shape-Name + Werte-Übersicht
'  fraFill           | Frame           | Caption = "Füllung"
'    chkFillColor    | CheckBox        | Caption = "Füllfarbe"
'  fraLine           | Frame           | Caption = "Linie"
'    chkLineColor    | CheckBox        | Caption = "Linienfarbe"
'    chkLineWeight   | CheckBox        | Caption = "Linienstärke"
'    chkLineDash     | CheckBox        | Caption = "Linienstil"
'  fraFont           | Frame           | Caption = "Schrift"
'    chkFontName     | CheckBox        | Caption = "Schriftart"
'    chkFontSize     | CheckBox        | Caption = "Schriftgröße"
'    chkFontBold     | CheckBox        | Caption = "Fett"
'    chkFontItalic   | CheckBox        | Caption = "Kursiv"
'    chkFontUnder    | CheckBox        | Caption = "Unterstrichen"
'    chkFontColor    | CheckBox        | Caption = "Schriftfarbe"
'  fraText           | Frame           | Caption = "Ausrichtung"
'    chkTextAlignH   | CheckBox        | Caption = "Horizontal"
'    chkTextAlignV   | CheckBox        | Caption = "Vertikal"
'  fraSize           | Frame           | Caption = "Größe"
'    chkWidth        | CheckBox        | Caption = "Breite"
'    chkHeight       | CheckBox        | Caption = "Höhe"
'  btnApply          | CommandButton   | Caption = "Anwenden"
'  btnSelectAll      | CommandButton   | Caption = "Alle"
'  btnNone           | CommandButton   | Caption = "Keine"
'  btnClose          | CommandButton   | Caption = "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Format Painter+"
'  Width:       320 pt,  Height: 420 pt
'  BorderStyle: 1 (Single)
' =============================================================================

Option Explicit


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form mit gecapturten Quellwerten.
'             Setzt Informations-Label und Checkbox-Defaults.
'             Muss vor Form.Show aufgerufen werden.
' -----------------------------------------------------------------------
Public Sub InitForm()

    On Error Resume Next

    ' --- Quell-Informationen anzeigen
    Dim snap As FormatSnapshot
    snap = modFormatPainterPlus.g_Snapshot

    Dim info As String
    info = "Quelle: " & snap.SourceName & vbCrLf

    ' Fill
    If snap.FillVisible And snap.FillType = msoFillSolid Then
        info = info & "Füllung: " & ColorToHex(snap.FillColor)
        If snap.FillTransp > 0 Then
            info = info & " (" & Format(snap.FillTransp * 100, "0") & "% transparent)"
        End If
        info = info & vbCrLf
    ElseIf Not snap.FillVisible Then
        info = info & "Füllung: keine" & vbCrLf
    End If

    ' Line
    If snap.LineVisible Then
        info = info & "Linie: " & ColorToHex(snap.LineColor) & _
               " | " & Format(snap.LineWeight, "0.0") & " pt" & vbCrLf
    End If

    ' Font
    If snap.FontName <> "" Then
        info = info & "Schrift: " & snap.FontName & " " & _
               Format(snap.FontSize, "0") & " pt"
        Dim attrs As String
        If snap.FontBold Then attrs = attrs & " B"
        If snap.FontItalic Then attrs = attrs & " I"
        If snap.FontUnderline Then attrs = attrs & " U"
        If Len(attrs) > 0 Then info = info & " [" & Trim(attrs) & "]"
        info = info & vbCrLf
    End If

    ' Size
    info = info & "Größe: " & Format(snap.ShapeWidth, "0.0") & _
           " × " & Format(snap.ShapeHeight, "0.0") & " pt"

    lblSourceInfo.Caption = info

    ' --- Checkboxen: alle standardmäßig aktiviert, außer Größe
    chkFillColor.Value = True
    chkLineColor.Value = True
    chkLineWeight.Value = True
    chkLineDash.Value = True
    chkFontName.Value = True
    chkFontSize.Value = True
    chkFontBold.Value = True
    chkFontItalic.Value = True
    chkFontUnder.Value = True
    chkFontColor.Value = True
    chkTextAlignH.Value = True
    chkTextAlignV.Value = True
    chkWidth.Value = False
    chkHeight.Value = False

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Liest Checkboxen aus, baut ApplyOptions und ruft Modul auf.
' -----------------------------------------------------------------------
Private Sub btnApply_Click()

    On Error GoTo ErrHandler

    Dim opts As modFormatPainterPlus.ApplyOptions

    opts.FillColor   = (chkFillColor.Value = True)
    opts.LineColor   = (chkLineColor.Value = True)
    opts.LineWeight  = (chkLineWeight.Value = True)
    opts.LineDash    = (chkLineDash.Value = True)
    opts.FontName    = (chkFontName.Value = True)
    opts.FontSize    = (chkFontSize.Value = True)
    opts.FontBold    = (chkFontBold.Value = True)
    opts.FontItalic  = (chkFontItalic.Value = True)
    opts.FontUnderline = (chkFontUnder.Value = True)
    opts.FontColor   = (chkFontColor.Value = True)
    opts.TextAlignH  = (chkTextAlignH.Value = True)
    opts.TextAlignV  = (chkTextAlignV.Value = True)
    opts.ShapeWidth  = (chkWidth.Value = True)
    opts.ShapeHeight = (chkHeight.Value = True)

    modFormatPainterPlus.ApplyFormatToSelection opts

    Exit Sub
ErrHandler:
    MsgBox "Fehler: " & Err.Description, vbExclamation, "Infront Toolkit – Format Painter+"
End Sub


' -----------------------------------------------------------------------
' Zweck:      Alle Checkboxen aktivieren.
' -----------------------------------------------------------------------
Private Sub btnSelectAll_Click()
    chkFillColor.Value  = True
    chkLineColor.Value  = True
    chkLineWeight.Value = True
    chkLineDash.Value   = True
    chkFontName.Value   = True
    chkFontSize.Value   = True
    chkFontBold.Value   = True
    chkFontItalic.Value = True
    chkFontUnder.Value  = True
    chkFontColor.Value  = True
    chkTextAlignH.Value = True
    chkTextAlignV.Value = True
    chkWidth.Value      = True
    chkHeight.Value     = True
End Sub


' -----------------------------------------------------------------------
' Zweck:      Alle Checkboxen deaktivieren.
' -----------------------------------------------------------------------
Private Sub btnNone_Click()
    chkFillColor.Value  = False
    chkLineColor.Value  = False
    chkLineWeight.Value = False
    chkLineDash.Value   = False
    chkFontName.Value   = False
    chkFontSize.Value   = False
    chkFontBold.Value   = False
    chkFontItalic.Value = False
    chkFontUnder.Value  = False
    chkFontColor.Value  = False
    chkTextAlignH.Value = False
    chkTextAlignV.Value = False
    chkWidth.Value      = False
    chkHeight.Value     = False
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


' -----------------------------------------------------------------------
' Zweck:      Hilfsfunktion – RGB-Long zu #RRGGBB-String.
' -----------------------------------------------------------------------
Private Function ColorToHex(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And &HFF
    g = (colorVal \ &H100) And &HFF
    b = (colorVal \ &H10000) And &HFF
    ColorToHex = "#" & Right("00" & Hex(r), 2) & _
                        Right("00" & Hex(g), 2) & _
                        Right("00" & Hex(b), 2)
End Function
