Attribute VB_Name = "frmFindReplace"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmFindReplace
' Zweck:  Globales Suchen & Ersetzen über alle / ausgewählte / aktuelle Folie(n).
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name               | Typ             | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lblFind            | Label           | Caption = "Suchen:"
'  txtFind            | TextBox         | Suchtext
'  lblReplace         | Label           | Caption = "Ersetzen durch:"
'  txtReplace         | TextBox         | Ersatztext
'  chkMatchCase       | CheckBox        | Caption = "Groß-/Kleinschreibung"
'  chkWholeWord       | CheckBox        | Caption = "Nur ganze Wörter"
'  chkIncludeNotes    | CheckBox        | Caption = "Notizen einschließen"
'  fraScope           | Frame           | Caption = "Bereich"
'    optAllSlides     | OptionButton    | Caption = "Alle Folien"
'    optSelectedSlides| OptionButton    | Caption = "Selektierte Folien"
'    optCurrentSlide  | OptionButton    | Caption = "Aktuelle Folie"
'  fraTargetShapes    | Frame           | Caption = "Shapes"
'    optAllShapes     | OptionButton    | Caption = "Alle"
'    optPlaceholders  | OptionButton    | Caption = "Nur Platzhalter/Titel"
'    optTextBoxes     | OptionButton    | Caption = "Nur Textboxen"
'  lblResult          | Label           | Zeigt Ergebnis / Trefferzahl an
'  btnPreview         | CommandButton   | Caption = "Vorschau (zählen)"
'  btnReplaceAll      | CommandButton   | Caption = "Alle ersetzen"
'  btnClose           | CommandButton   | Caption = "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Find & Replace"
'  Width:       310 pt,  Height: 370 pt
'  BorderStyle: 1 (Single)
'
' Hinweis Formatierungserhalt:
'  Der Ersatz erfolgt run-weise. Treffer die über Run-Grenzen gehen
'  (z.B. "Hallo" wenn "H" fett und "allo" normal) werden nicht ersetzt,
'  da das Aufteilen von Formatierungsgrenzen zu undefinierten Ergebnissen
'  führen würde.
' =============================================================================

Option Explicit


' -----------------------------------------------------------------------
' Zweck:      Initialisiert die Form – setzt Standardwerte.
'             Wird beim ersten Öffnen (UserForm_Initialize) aufgerufen.
' -----------------------------------------------------------------------
Private Sub UserForm_Initialize()

    ' Scope-Defaults
    optAllSlides.Value = True

    ' Shape-Filter-Default
    optAllShapes.Value = True

    ' Optionen-Defaults
    chkMatchCase.Value = False
    chkWholeWord.Value = False
    chkIncludeNotes.Value = False

    lblResult.Caption = ""

    txtFind.SetFocus
End Sub


' -----------------------------------------------------------------------
' Zweck:      Zählt Treffer und zeigt Ergebnis an (ohne zu ersetzen).
' -----------------------------------------------------------------------
Private Sub btnPreview_Click()

    On Error GoTo ErrHandler

    If txtFind.Text = "" Then
        lblResult.Caption = "Bitte Suchtext eingeben."
        Exit Sub
    End If

    Dim opts As modFindReplace.FindReplaceOptions
    BuildOptions opts

    Dim count As Long
    count = modFindReplace.CountMatches(opts)

    If count = 0 Then
        lblResult.Caption = "Keine Treffer gefunden."
    Else
        lblResult.Caption = count & " Treffer gefunden."
    End If

    Exit Sub
ErrHandler:
    lblResult.Caption = "Fehler: " & Err.Description
End Sub


' -----------------------------------------------------------------------
' Zweck:      Ersetzt alle Treffer und zeigt Ergebnis an.
' -----------------------------------------------------------------------
Private Sub btnReplaceAll_Click()

    On Error GoTo ErrHandler

    If txtFind.Text = "" Then
        lblResult.Caption = "Bitte Suchtext eingeben."
        Exit Sub
    End If

    Dim opts As modFindReplace.FindReplaceOptions
    BuildOptions opts

    ' Bestätigung bei "Alle Folien"
    If opts.Scope = 0 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("In allen Folien ersetzen?" & vbCrLf & vbCrLf & _
                        "Suchen:   """ & opts.FindText & """" & vbCrLf & _
                        "Ersetzen: """ & opts.ReplaceText & """", _
                        vbQuestion + vbOKCancel, "Infront Toolkit – Find & Replace")
        If answer <> vbOK Then Exit Sub
    End If

    lblResult.Caption = "Wird ausgeführt..."
    DoEvents

    Dim count As Long
    count = modFindReplace.ExecuteReplace(opts)

    If count = 0 Then
        lblResult.Caption = "Keine Treffer gefunden."
    Else
        lblResult.Caption = count & " Ersetzung" & IIf(count = 1, "", "en") & " vorgenommen."
    End If

    Exit Sub
ErrHandler:
    lblResult.Caption = "Fehler: " & Err.Description
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
' Zweck:      Enter in txtFind springt zu txtReplace.
' -----------------------------------------------------------------------
Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then   ' Enter
        txtReplace.SetFocus
        KeyCode = 0
    End If
End Sub


' -----------------------------------------------------------------------
' Zweck:      Enter in txtReplace löst Ersetzen aus.
' -----------------------------------------------------------------------
Private Sub txtReplace_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then   ' Enter
        btnReplaceAll_Click
        KeyCode = 0
    End If
End Sub


' -----------------------------------------------------------------------
' Zweck:      Liest Form-Controls aus und befüllt FindReplaceOptions.
' Parameter:  opts - Ausgabe
' -----------------------------------------------------------------------
Private Sub BuildOptions(ByRef opts As modFindReplace.FindReplaceOptions)

    opts.FindText    = txtFind.Text
    opts.ReplaceText = txtReplace.Text
    opts.MatchCase   = (chkMatchCase.Value = True)
    opts.WholeWord   = (chkWholeWord.Value = True)
    opts.IncludeNotes = (chkIncludeNotes.Value = True)

    ' Scope
    If optAllSlides.Value Then
        opts.Scope = 0
    ElseIf optSelectedSlides.Value Then
        opts.Scope = 1
    Else
        opts.Scope = 2
    End If

    ' Shape-Filter
    If optAllShapes.Value Then
        opts.TargetShapes = 0
    ElseIf optPlaceholders.Value Then
        opts.TargetShapes = 1
    Else
        opts.TargetShapes = 2
    End If
End Sub
