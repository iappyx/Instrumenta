Attribute VB_Name = "modFindReplace"
Option Explicit

' =============================================================================
' Modul:  modFindReplace
' Zweck:  Globales Suchen & Ersetzen über alle / ausgewählte / aktuelle Folie(n).
'         Traversiert Gruppen rekursiv, Tabellen zellenweise, TextFrames run-weise.
'         Zeichenformatierung (Fett, Kursiv, Farbe usw.) der Runs bleibt erhalten.
'
' Unterschied zu eingebautem ReplaceDialog (Ctrl+H):
'   - Scope wählbar: Alle / Selektierte / Aktuelle Folie
'   - Selektiver Shape-Filter: Alle / Nur Platzhalter / Nur Textboxen
'   - Optional Sprechernotizen einschließen
'   - Run-genaues Ersetzen: Formatierung bleibt je Run erhalten
'
' Plattform:  Windows und Mac
' =============================================================================

' --- Öffentlicher Options-Typ (von frmFindReplace befüllt) -------------------

Public Type FindReplaceOptions
    FindText        As String
    ReplaceText     As String
    MatchCase       As Boolean
    WholeWord       As Boolean
    ' Scope: 0 = Alle Folien, 1 = Selektierte Folien, 2 = Aktuelle Folie
    Scope           As Long
    IncludeNotes    As Boolean
    ' TargetShapes: 0 = Alle, 1 = Nur Platzhalter/Titel, 2 = Nur Textboxen
    TargetShapes    As Long
End Type


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – öffnet die Find & Replace Form.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub ShowFindReplace(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Find & Replace"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    frmFindReplace.Show vbModeless

    Exit Sub
ErrHandler:
    MsgBox "Fehler in ShowFindReplace: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Führt Suchen & Ersetzen durch.
' Parameter:  opts - FindReplaceOptions
' Rückgabe:   Anzahl der vorgenommenen Ersetzungen
' -----------------------------------------------------------------------
Public Function ExecuteReplace(opts As FindReplaceOptions) As Long

    On Error GoTo ErrHandler

    If opts.FindText = "" Then
        ExecuteReplace = 0
        Exit Function
    End If

    Dim slides As SlideRange
    Set slides = GetScopeSlides(opts.Scope)
    If slides Is Nothing Then
        ExecuteReplace = 0
        Exit Function
    End If

    Dim total As Long
    total = 0

    Dim sld As Slide
    For Each sld In slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            total = total + ReplaceInShape(shp, opts)
        Next shp

        ' Sprechernotizen
        If opts.IncludeNotes Then
            On Error Resume Next
            If sld.HasNotesPage Then
                Dim notePage As Slide
                Set notePage = sld.NotesPage
                Dim nShp As Shape
                For Each nShp In notePage.Shapes
                    If nShp.HasTextFrame Then
                        total = total + ReplaceInTextRange(nShp.TextFrame.TextRange, opts)
                    End If
                Next nShp
            End If
            On Error GoTo ErrHandler
        End If
    Next sld

    ExecuteReplace = total
    Exit Function

ErrHandler:
    ExecuteReplace = total
End Function


' -----------------------------------------------------------------------
' Zweck:      Zählt Treffer ohne zu ersetzen (Preview).
' Parameter:  opts - FindReplaceOptions
' Rückgabe:   Anzahl der Treffer
' -----------------------------------------------------------------------
Public Function CountMatches(opts As FindReplaceOptions) As Long

    On Error GoTo ErrHandler

    If opts.FindText = "" Then
        CountMatches = 0
        Exit Function
    End If

    Dim slides As SlideRange
    Set slides = GetScopeSlides(opts.Scope)
    If slides Is Nothing Then
        CountMatches = 0
        Exit Function
    End If

    Dim total As Long
    total = 0

    Dim sld As Slide
    For Each sld In slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            total = total + CountInShape(shp, opts)
        Next shp

        If opts.IncludeNotes Then
            On Error Resume Next
            If sld.HasNotesPage Then
                Dim notePage As Slide
                Set notePage = sld.NotesPage
                Dim nShp As Shape
                For Each nShp In notePage.Shapes
                    If nShp.HasTextFrame Then
                        total = total + CountInTextRange( _
                            nShp.TextFrame.TextRange, opts)
                    End If
                Next nShp
            End If
            On Error GoTo ErrHandler
        End If
    Next sld

    CountMatches = total
    Exit Function

ErrHandler:
    CountMatches = total
End Function


' =============================================================================
' SHAPE-TRAVERSIERUNG
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ersetzt Text in einem Shape (Gruppe/Tabelle/TextFrame).
' Rückgabe:   Anzahl Ersetzungen
' -----------------------------------------------------------------------
Private Function ReplaceInShape(shp As Shape, opts As FindReplaceOptions) As Long

    On Error Resume Next

    Dim count As Long
    count = 0

    ' Shape-Typ-Filter prüfen
    If Not ShapeMatchesTarget(shp, opts.TargetShapes) Then
        ReplaceInShape = 0
        Exit Function
    End If

    ' Gruppe → rekursiv
    If shp.Type = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            count = count + ReplaceInShape(gi, opts)
        Next gi
        ReplaceInShape = count
        Exit Function
    End If

    ' Tabelle → zellenweise
    If shp.HasTable Then
        Dim tbl As Table
        Set tbl = shp.Table
        If Not tbl Is Nothing Then
            Dim r As Long, c As Long
            For r = 1 To tbl.Rows.Count
                For c = 1 To tbl.Columns.Count
                    Dim cel As cell
                    Set cel = tbl.cell(r, c)
                    If cel.Shape.HasTextFrame Then
                        count = count + ReplaceInTextRange( _
                            cel.Shape.TextFrame.TextRange, opts)
                    End If
                Next c
            Next r
        End If
        ReplaceInShape = count
        Exit Function
    End If

    ' Normales Shape mit TextFrame
    If shp.HasTextFrame Then
        count = count + ReplaceInTextRange(shp.TextFrame.TextRange, opts)
    End If

    ReplaceInShape = count
End Function


' -----------------------------------------------------------------------
' Zweck:      Zählt Treffer in einem Shape ohne zu ersetzen.
' Rückgabe:   Anzahl Treffer
' -----------------------------------------------------------------------
Private Function CountInShape(shp As Shape, opts As FindReplaceOptions) As Long

    On Error Resume Next

    Dim count As Long
    count = 0

    If Not ShapeMatchesTarget(shp, opts.TargetShapes) Then
        CountInShape = 0
        Exit Function
    End If

    If shp.Type = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            count = count + CountInShape(gi, opts)
        Next gi
        CountInShape = count
        Exit Function
    End If

    If shp.HasTable Then
        Dim tbl As Table
        Set tbl = shp.Table
        If Not tbl Is Nothing Then
            Dim r As Long, c As Long
            For r = 1 To tbl.Rows.Count
                For c = 1 To tbl.Columns.Count
                    Dim cel As cell
                    Set cel = tbl.cell(r, c)
                    If cel.Shape.HasTextFrame Then
                        count = count + CountInTextRange( _
                            cel.Shape.TextFrame.TextRange, opts)
                    End If
                Next c
            Next r
        End If
        CountInShape = count
        Exit Function
    End If

    If shp.HasTextFrame Then
        count = count + CountInTextRange(shp.TextFrame.TextRange, opts)
    End If

    CountInShape = count
End Function


' -----------------------------------------------------------------------
' Zweck:      Ersetzt Text run-weise in einem TextRange.
'             Jeder Run wird einzeln behandelt → Formatierung bleibt erhalten.
'             Treffer die über Run-Grenzen gehen werden nicht ersetzt
'             (da Formatierungs-Bruch unklar; dokumentiert).
' Rückgabe:   Anzahl Ersetzungen
' -----------------------------------------------------------------------
Public Function ReplaceInTextRange(tr As TextRange, _
                                   opts As FindReplaceOptions) As Long
    On Error Resume Next

    Dim count As Long
    count = 0

    Dim p As Long, rn As Long
    For p = 1 To tr.Paragraphs.Count
        For rn = 1 To tr.Paragraphs(p).Runs.Count
            Dim run As TextRange
            Set run = tr.Paragraphs(p).Runs(rn)

            Dim original As String
            original = run.Text

            Dim replaced As String
            Dim replCount As Long
            replaced = ReplaceString(original, opts.FindText, opts.ReplaceText, _
                                     opts.MatchCase, opts.WholeWord, replCount)

            If replCount > 0 Then
                run.Text = replaced
                count = count + replCount
            End If
        Next rn
    Next p

    On Error GoTo 0
    ReplaceInTextRange = count
End Function


' -----------------------------------------------------------------------
' Zweck:      Zählt Treffer run-weise ohne zu ersetzen.
' Rückgabe:   Anzahl Treffer
' -----------------------------------------------------------------------
Private Function CountInTextRange(tr As TextRange, _
                                  opts As FindReplaceOptions) As Long
    On Error Resume Next

    Dim count As Long
    count = 0

    Dim p As Long, rn As Long
    For p = 1 To tr.Paragraphs.Count
        For rn = 1 To tr.Paragraphs(p).Runs.Count
            Dim run As TextRange
            Set run = tr.Paragraphs(p).Runs(rn)
            Dim dummy As Long
            Call ReplaceString(run.Text, opts.FindText, opts.ReplaceText, _
                               opts.MatchCase, opts.WholeWord, dummy)
            count = count + dummy
        Next rn
    Next p

    On Error GoTo 0
    CountInTextRange = count
End Function


' =============================================================================
' HILFS-FUNKTIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Sucht und ersetzt in einem String, respektiert MatchCase
'             und WholeWord. Gibt Anzahl Ersetzungen per Ref zurück.
' Parameter:  src        - Quellstring
'             findStr    - Suchstring
'             replStr    - Ersatzstring
'             matchCase  - Groß-/Kleinschreibung beachten
'             wholeWord  - Nur ganze Wörter
'             replCount  - Ausgabe: Anzahl Ersetzungen
' Rückgabe:   Neuer String
' -----------------------------------------------------------------------
Private Function ReplaceString(src As String, findStr As String, _
                               replStr As String, matchCase As Boolean, _
                               wholeWord As Boolean, _
                               ByRef replCount As Long) As String
    replCount = 0

    If Len(findStr) = 0 Or Len(src) = 0 Then
        ReplaceString = src
        Exit Function
    End If

    Dim compareSrc As String
    Dim compareFind As String

    If matchCase Then
        compareSrc = src
        compareFind = findStr
    Else
        compareSrc = LCase(src)
        compareFind = LCase(findStr)
    End If

    Dim result As String
    result = ""
    Dim pos As Long
    Dim lastPos As Long
    lastPos = 1

    Do
        pos = InStr(lastPos, compareSrc, compareFind)
        If pos = 0 Then Exit Do

        ' WholeWord-Prüfung
        If wholeWord Then
            Dim charBefore As String
            Dim charAfter As String
            If pos > 1 Then
                charBefore = Mid(src, pos - 1, 1)
            Else
                charBefore = " "
            End If
            If pos + Len(findStr) <= Len(src) Then
                charAfter = Mid(src, pos + Len(findStr), 1)
            Else
                charAfter = " "
            End If

            If IsWordChar(charBefore) Or IsWordChar(charAfter) Then
                ' Kein ganzes Wort – Zeichen übernehmen und weiter
                result = result & Mid(src, lastPos, pos - lastPos + 1)
                lastPos = pos + 1
                GoTo NextMatch
            End If
        End If

        ' Treffer
        result = result & Mid(src, lastPos, pos - lastPos) & replStr
        replCount = replCount + 1
        lastPos = pos + Len(findStr)

NextMatch:
    Loop

    result = result & Mid(src, lastPos)
    ReplaceString = result
End Function


' -----------------------------------------------------------------------
' Zweck:      Prüft ob ein Zeichen ein Wortzeichen ist (für WholeWord).
' -----------------------------------------------------------------------
Private Function IsWordChar(c As String) As Boolean
    If Len(c) = 0 Then
        IsWordChar = False
        Exit Function
    End If
    Dim code As Long
    code = Asc(c)
    IsWordChar = (code >= 65 And code <= 90) Or _   ' A-Z
                 (code >= 97 And code <= 122) Or _  ' a-z
                 (code >= 48 And code <= 57) Or _   ' 0-9
                 (code = 95)                         ' _
End Function


' -----------------------------------------------------------------------
' Zweck:      Ermittelt SlideRange anhand des Scope-Werts.
' Parameter:  scope - 0=Alle, 1=Selektierte, 2=Aktuelle
' Rückgabe:   SlideRange oder Nothing bei Fehler
' -----------------------------------------------------------------------
Private Function GetScopeSlides(scope As Long) As SlideRange

    On Error GoTo ErrHandler

    Select Case scope
        Case 0  ' Alle Folien
            Set GetScopeSlides = ActivePresentation.Slides.Range

        Case 1  ' Selektierte Folien
            Dim sel As Selection
            Set sel = ActiveWindow.Selection
            If sel.Type = ppSelectionSlides Then
                Set GetScopeSlides = sel.SlideRange
            Else
                ' Fallback: aktuelle Folie
                Set GetScopeSlides = ActivePresentation.Slides.Range( _
                    Array(ActiveWindow.View.Slide.SlideIndex))
            End If

        Case 2  ' Aktuelle Folie
            Set GetScopeSlides = ActivePresentation.Slides.Range( _
                Array(ActiveWindow.View.Slide.SlideIndex))

        Case Else
            Set GetScopeSlides = ActivePresentation.Slides.Range
    End Select

    Exit Function
ErrHandler:
    Set GetScopeSlides = Nothing
End Function


' -----------------------------------------------------------------------
' Zweck:      Prüft ob ein Shape dem TargetShapes-Filter entspricht.
' Parameter:  shp          - zu prüfendes Shape
'             targetFilter - 0=Alle, 1=Platzhalter/Titel, 2=Textboxen
' Rückgabe:   True wenn Shape geprüft werden soll
' -----------------------------------------------------------------------
Private Function ShapeMatchesTarget(shp As Shape, targetFilter As Long) As Boolean

    On Error Resume Next

    Select Case targetFilter
        Case 0  ' Alle Shapes mit Text
            ShapeMatchesTarget = True

        Case 1  ' Nur Platzhalter und Titel
            Dim ptype As Long
            ptype = shp.PlaceholderFormat.Type
            ShapeMatchesTarget = (Err.Number = 0)  ' hat PlaceholderFormat → ist Platzhalter
            Err.Clear

        Case 2  ' Nur Textboxen (msoTextBox)
            ShapeMatchesTarget = (shp.Type = msoTextBox)

        Case Else
            ShapeMatchesTarget = True
    End Select

    On Error GoTo 0
End Function
