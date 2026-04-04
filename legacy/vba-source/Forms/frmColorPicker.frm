Attribute VB_Name = "frmColorPicker"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmColorPicker
' Zweck:  Ergebnisform für den Screen Color Picker.
'         Zeigt Farbvorschau, Hex-Code und RGB-Werte.
'         Ermöglicht Anwenden der Farbe auf Fill / Line / Font.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name           | Typ              | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lblPreview     | Label            | BackColor = aufgenommene Farbe (Vorschau)
'  lblInfo        | Label            | Zeigt "#RRGGBB" und "R: x G: x B: x"
'  optFill        | OptionButton     | "Füllfarbe"   (Value = True als Default)
'  optLine        | OptionButton     | "Linienfarbe"
'  optFont        | OptionButton     | "Schriftfarbe"
'  btnApply       | CommandButton    | "Anwenden"
'  btnClose       | CommandButton    | "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:   "Infront Toolkit – Color Picker"
'  Width:     240 pt,  Height: 200 pt
'  BorderStyle: 1 (Single)
' =============================================================================

Option Explicit

Private m_PickedColor As Long


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form mit der aufgenommenen Farbe.
'             Muss vor Form.Show aufgerufen werden.
' Parameter:  pickedColor - aufgenommener RGB-Farbwert
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub InitForm(pickedColor As Long)

    m_PickedColor = pickedColor

    On Error Resume Next   ' Controls möglicherweise noch nicht vorhanden

    ' Farbvorschau
    lblPreview.BackColor = pickedColor
    lblPreview.BorderStyle = fmBorderStyleSingle

    ' Hex + RGB Text
    lblInfo.Caption = modColorPicker.ColorToHex(pickedColor) & vbCrLf & _
                      modColorPicker.ColorToRGBString(pickedColor)

    ' Standard-Option: Füllfarbe
    optFill.Value = True

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Wendet die aufgenommene Farbe auf die Selektion an und
'             schließt die Form.
' Parameter:  (keiner)
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub btnApply_Click()

    Dim target As String

    On Error Resume Next
    If optFill.Value Then
        target = "fill"
    ElseIf optLine.Value Then
        target = "line"
    Else
        target = "font"
    End If
    On Error GoTo 0

    If target = "" Then target = "fill"

    Me.Hide
    modColorPicker.ApplyColorToSelection m_PickedColor, target
    Unload Me
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form ohne Änderungen.
' Parameter:  (keiner)
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form sauber bei X-Button oder Alt+F4.
' Parameter:  Cancel - kann auf True gesetzt werden um Schließen zu verhindern
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = False   ' Schließen erlauben
    End If
End Sub
