Attribute VB_Name = "modGapEqualizer"
Option Explicit

' =============================================================================
' Modul:  modGapEqualizer
' Zweck:  Smart Gap Equalizer – setzt exakte Abstände zwischen Shapes.
'
' Konzept:
'   - Gap = Abstand zwischen rechter Kante Shape[i] und linker Kante Shape[i+1]
'     (horizontal) bzw. unterer Kante und oberer Kante (vertikal)
'   - Negative Gaps (Überlappungen) sind erlaubt und werden korrekt dargestellt
'   - Mindestens 2 Shapes für sinnvolle Anwendung
'   - Shapes werden vor der Positionierung nach Left (H) oder Top (V) sortiert
'
' Plattform:  Windows und Mac
' =============================================================================

' --- Öffentlicher Options-Typ (von frmGapEqualizer befüllt) ------------------

Public Type GapOptions
    Horizontal    As Boolean
    Vertical      As Boolean
    ' GapMode: 0=Custom, 1=Average, 2=Minimum, 3=Maximum
    GapMode       As Long
    CustomGapPt   As Single    ' nur bei GapMode=0
    ' AnchorMode: 0=Erstes Shape fixiert, 1=Innerhalb Bounds verteilen
    AnchorMode    As Long
End Type


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – öffnet den Gap Equalizer.
' -----------------------------------------------------------------------
Public Sub ShowGapEqualizer(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Gap Equalizer"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Bitte mindestens 2 Shapes selektieren.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Bitte mindestens 2 Shapes selektieren.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    frmGapEqualizer.InitForm
    frmGapEqualizer.Show vbModeless

    Exit Sub
ErrHandler:
    MsgBox "Fehler in ShowGapEqualizer: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Führt Gap-Equalisierung durch.
' Parameter:  opts - GapOptions
' -----------------------------------------------------------------------
Public Sub EqualizeGaps(opts As GapOptions)

    Const DLG_TITLE As String = "Infront Toolkit – Gap Equalizer"

    On Error GoTo ErrHandler

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Bitte Shapes selektieren.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Mindestens 2 Shapes erforderlich.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim changed As Boolean
    changed = False

    If opts.Horizontal Then
        Dim hShapes() As Shape
        GetShapesSorted sr, True, hShapes

        Dim targetH As Single
        targetH = ResolveTargetGap(hShapes, True, opts)

        ApplyGaps hShapes, targetH, True, opts.AnchorMode
        changed = True
    End If

    If opts.Vertical Then
        Dim vShapes() As Shape
        GetShapesSorted sr, False, vShapes

        Dim targetV As Single
        targetV = ResolveTargetGap(vShapes, False, opts)

        ApplyGaps vShapes, targetV, False, opts.AnchorMode
        changed = True
    End If

    If Not changed Then
        MsgBox "Bitte mindestens eine Richtung (H oder V) wählen.", _
               vbExclamation, DLG_TITLE
    End If

    Exit Sub
ErrHandler:
    MsgBox "Fehler in EqualizeGaps: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Berechnet aktuelle Gap-Informationen (für Preview in Form).
' Parameter:  isHorizontal - True=H, False=V
' Rückgabe:   String mit Min/Max/Avg-Angabe in pt
' -----------------------------------------------------------------------
Public Function GetGapInfo(isHorizontal As Boolean) As String

    On Error GoTo ReturnEmpty

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then GoTo ReturnEmpty
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then GoTo ReturnEmpty

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    Dim shapes() As Shape
    GetShapesSorted sr, isHorizontal, shapes

    Dim gaps() As Single
    CalculateCurrentGaps shapes, isHorizontal, gaps

    Dim n As Long
    n = UBound(gaps)
    If n < 1 Then GoTo ReturnEmpty

    Dim minG As Single, maxG As Single, sumG As Single
    minG = gaps(1): maxG = gaps(1): sumG = 0

    Dim i As Long
    For i = 1 To n
        If gaps(i) < minG Then minG = gaps(i)
        If gaps(i) > maxG Then maxG = gaps(i)
        sumG = sumG + gaps(i)
    Next i

    Dim avg As Single
    avg = sumG / n

    Dim dir As String
    dir = IIf(isHorizontal, "H", "V")

    GetGapInfo = dir & ": Ø" & Format(avg, "0.0") & " pt" & _
                 "  Min " & Format(minG, "0.0") & _
                 "  Max " & Format(maxG, "0.0")
    Exit Function

ReturnEmpty:
    GetGapInfo = IIf(isHorizontal, "H", "V") & ": –"
End Function


' =============================================================================
' SORT + GAP-BERECHNUNG
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Füllt shapes()-Array mit Shapes aus ShapeRange, sortiert nach
'             Left (horizontal) oder Top (vertikal). Bubble-Sort.
' Parameter:  sr           - Quell-ShapeRange
'             isHorizontal - True=nach Left sortieren, False=nach Top
'             shapes()     - Ausgabe-Array (1-basiert)
' -----------------------------------------------------------------------
Public Sub GetShapesSorted(sr As ShapeRange, isHorizontal As Boolean, _
                           ByRef shapes() As Shape)
    Dim n As Long
    n = sr.Count
    ReDim shapes(1 To n)

    Dim i As Long
    For i = 1 To n
        Set shapes(i) = sr(i)
    Next i

    ' Bubble-Sort
    Dim j As Long
    Dim swapped As Boolean
    Dim tempShp As Shape

    Do
        swapped = False
        For i = 1 To n - 1
            Dim keyA As Single
            Dim keyB As Single
            If isHorizontal Then
                keyA = shapes(i).Left
                keyB = shapes(i + 1).Left
            Else
                keyA = shapes(i).Top
                keyB = shapes(i + 1).Top
            End If
            If keyA > keyB Then
                Set tempShp = shapes(i)
                Set shapes(i) = shapes(i + 1)
                Set shapes(i + 1) = tempShp
                swapped = True
            End If
        Next i
    Loop While swapped
End Sub


' -----------------------------------------------------------------------
' Zweck:      Berechnet die Gaps zwischen aufeinanderfolgenden Shapes.
' Parameter:  shapes()     - sortiertes Shape-Array (1-basiert)
'             isHorizontal - True=H, False=V
'             gaps()       - Ausgabe: Array der Gaps (1-basiert, n-1 Einträge)
' -----------------------------------------------------------------------
Public Sub CalculateCurrentGaps(shapes() As Shape, isHorizontal As Boolean, _
                                ByRef gaps() As Single)
    Dim n As Long
    n = UBound(shapes)

    If n < 2 Then
        ReDim gaps(0)
        Exit Sub
    End If

    ReDim gaps(1 To n - 1)

    Dim i As Long
    For i = 1 To n - 1
        If isHorizontal Then
            gaps(i) = shapes(i + 1).Left - (shapes(i).Left + shapes(i).Width)
        Else
            gaps(i) = shapes(i + 1).Top - (shapes(i).Top + shapes(i).Height)
        End If
    Next i
End Sub


' -----------------------------------------------------------------------
' Zweck:      Bestimmt den Ziel-Gap-Wert anhand GapMode.
' Parameter:  shapes()     - sortiertes Shape-Array
'             isHorizontal - Richtung
'             opts         - GapOptions
' Rückgabe:   Ziel-Gap in pt
' -----------------------------------------------------------------------
Private Function ResolveTargetGap(shapes() As Shape, isHorizontal As Boolean, _
                                  opts As GapOptions) As Single

    If opts.GapMode = 0 Then
        ResolveTargetGap = opts.CustomGapPt
        Exit Function
    End If

    Dim gaps() As Single
    CalculateCurrentGaps shapes, isHorizontal, gaps

    Dim n As Long
    n = UBound(gaps)

    If n < 1 Then
        ResolveTargetGap = 0
        Exit Function
    End If

    Dim minG As Single, maxG As Single, sumG As Single
    minG = gaps(1): maxG = gaps(1): sumG = 0

    Dim i As Long
    For i = 1 To n
        If gaps(i) < minG Then minG = gaps(i)
        If gaps(i) > maxG Then maxG = gaps(i)
        sumG = sumG + gaps(i)
    Next i

    Select Case opts.GapMode
        Case 1: ResolveTargetGap = sumG / n   ' Average
        Case 2: ResolveTargetGap = minG        ' Minimum
        Case 3: ResolveTargetGap = maxG        ' Maximum
        Case Else: ResolveTargetGap = sumG / n
    End Select
End Function


' -----------------------------------------------------------------------
' Zweck:      Positioniert Shapes mit gleichem Gap.
' Parameter:  shapes()     - sortiertes Shape-Array (1-basiert)
'             targetGap    - Ziel-Gap in pt
'             isHorizontal - True=H (Left), False=V (Top)
'             anchorMode   - 0=Erstes Shape fixiert, 1=Bounds beibehalten
' -----------------------------------------------------------------------
Public Sub ApplyGaps(shapes() As Shape, targetGap As Single, _
                     isHorizontal As Boolean, anchorMode As Long)

    On Error Resume Next

    Dim n As Long
    n = UBound(shapes)
    If n < 2 Then Exit Sub

    If anchorMode = 0 Then
        ' Erstes Shape bleibt; jedes folgende Shape wird relativ positioniert
        Dim i As Long
        For i = 2 To n
            If isHorizontal Then
                shapes(i).Left = shapes(i - 1).Left + shapes(i - 1).Width + targetGap
            Else
                shapes(i).Top = shapes(i - 1).Top + shapes(i - 1).Height + targetGap
            End If
        Next i

    Else
        ' AnchorMode=1: Gesamtbreite/-höhe der Gruppe beibehalten,
        ' Shapes gleichmäßig verteilen (analog zu PPT Distribute)
        Dim totalSize As Single
        totalSize = 0

        Dim j As Long
        For j = 1 To n
            If isHorizontal Then
                totalSize = totalSize + shapes(j).Width
            Else
                totalSize = totalSize + shapes(j).Height
            End If
        Next j

        ' Bounds (Anfang erstes Shape bis Ende letztes Shape) beibehalten
        Dim startPos As Single
        Dim endPos As Single
        If isHorizontal Then
            startPos = shapes(1).Left
            endPos   = shapes(n).Left + shapes(n).Width
        Else
            startPos = shapes(1).Top
            endPos   = shapes(n).Top + shapes(n).Height
        End If

        Dim totalBounds As Single
        totalBounds = endPos - startPos

        ' Gap = (Bounds - Summe aller Shape-Größen) / (n-1)
        Dim equalGap As Single
        If n > 1 Then
            equalGap = (totalBounds - totalSize) / (n - 1)
        Else
            equalGap = 0
        End If

        ' Shapes neu positionieren
        Dim curPos As Single
        curPos = startPos
        For j = 1 To n
            If isHorizontal Then
                shapes(j).Left = curPos
                curPos = curPos + shapes(j).Width + equalGap
            Else
                shapes(j).Top = curPos
                curPos = curPos + shapes(j).Height + equalGap
            End If
        Next j
    End If

    On Error GoTo 0
End Sub
