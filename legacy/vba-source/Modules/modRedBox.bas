Attribute VB_Name = "modRedBox"
Option Explicit

' =============================================================================
' Modul:  modRedBox
' Zweck:  Red Box – fügt Hervorhebungsrahmen/-fläche auf Folien ein.
'
' Zwei Varianten:
'   Outline:  Transparenter Hintergrund, roter Rahmen (2.5 pt)
'   Filled:   Halbtransparentes rotes Fill (80% Transparenz) + roter Rahmen
'
' Positionierung:
'   Shapes selektiert → umrahmt Gesamtbounds + PADDING
'   Nichts selektiert → Folienmitte (DEFAULT_W × DEFAULT_H)
'
' Erkennung / Entfernen:
'   Tag InfrontRedBox=1 auf jedem eingefügten Shape
'
' Plattform:  Windows und Mac
' =============================================================================

Private Const REDBOX_TAG_KEY   As String = "InfrontRedBox"
Private Const REDBOX_TAG_VALUE As String = "1"
' Farbe: RGB(204, 0, 0) direkt in den Funktionen verwendet (kein Literal nötig)
Private Const REDBOX_LINE_WEIGHT As Single = 2.5
Private Const REDBOX_PADDING     As Single = 6    ' pt Abstand um Selektion
Private Const DEFAULT_W          As Single = 200  ' pt Breite wenn nichts selektiert
Private Const DEFAULT_H          As Single = 120  ' pt Höhe wenn nichts selektiert
Private Const FILL_TRANSPARENCY  As Single = 0.8  ' 80% transparent


' =============================================================================
' RIBBON-CALLBACKS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Fügt Red Box (Outline) ein.
' -----------------------------------------------------------------------
Public Sub InsertRedBox(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Red Box"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide

    Dim boxLeft As Single, boxTop As Single
    Dim boxW As Single, boxH As Single

    If HasShapesSelected() Then
        GetSelectionBounds REDBOX_PADDING, boxLeft, boxTop, boxW, boxH
    Else
        CenterBoxOnSlide boxLeft, boxTop, boxW, boxH
    End If

    Dim shp As Shape
    Set shp = CreateRedBoxShape(sld, boxLeft, boxTop, boxW, boxH, False)

    If Not shp Is Nothing Then
        shp.Select msoFalse  ' zur Folie hinzufügen ohne Selektion zu verlieren
    End If

    Exit Sub
ErrHandler:
    MsgBox "Fehler in InsertRedBox: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Fügt Red Box (Filled) ein.
' -----------------------------------------------------------------------
Public Sub InsertFilledRedBox(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Red Box (Filled)"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide

    Dim boxLeft As Single, boxTop As Single
    Dim boxW As Single, boxH As Single

    If HasShapesSelected() Then
        GetSelectionBounds REDBOX_PADDING, boxLeft, boxTop, boxW, boxH
    Else
        CenterBoxOnSlide boxLeft, boxTop, boxW, boxH
    End If

    Dim shp As Shape
    Set shp = CreateRedBoxShape(sld, boxLeft, boxTop, boxW, boxH, True)

    If Not shp Is Nothing Then
        shp.Select msoFalse
    End If

    Exit Sub
ErrHandler:
    MsgBox "Fehler in InsertFilledRedBox: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Entfernt alle Red Boxes von der aktuellen Folie.
' -----------------------------------------------------------------------
Public Sub RemoveRedBoxes(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Red Box"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide

    ' Zählen
    Dim count As Long
    count = CountRedBoxesOnSlide(sld)

    If count = 0 Then
        MsgBox "Keine Red Boxes auf dieser Folie.", vbInformation, DLG_TITLE
        Exit Sub
    End If

    ' Bestätigung nur bei mehr als 3
    If count > 3 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox(count & " Red Boxes entfernen?", _
                        vbQuestion + vbOKCancel, DLG_TITLE)
        If answer <> vbOK Then Exit Sub
    End If

    Dim removed As Long
    removed = DeleteRedBoxesFromSlide(sld)

    MsgBox removed & " Red Box" & IIf(removed = 1, "", "es") & " entfernt.", _
           vbInformation, DLG_TITLE

    Exit Sub
ErrHandler:
    MsgBox "Fehler in RemoveRedBoxes: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' =============================================================================
' SHAPE-ERSTELLUNG
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Erstellt das Red Box Shape auf der angegebenen Folie.
' Parameter:  sld    - Zielfolie
'             left, top, w, h - Position und Größe in pt
'             filled - True = halbtransparentes Fill, False = nur Outline
' Rückgabe:   Erstelltes Shape
' -----------------------------------------------------------------------
Public Function CreateRedBoxShape(sld As Slide, _
                                  left As Single, top As Single, _
                                  w As Single, h As Single, _
                                  filled As Boolean) As Shape

    On Error GoTo ErrHandler

    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, left, top, w, h)

    ' Tag setzen
    shp.Tags.Add REDBOX_TAG_KEY, REDBOX_TAG_VALUE

    ' Rahmen
    With shp.Line
        .Visible    = msoTrue
        .ForeColor.RGB = RGB(204, 0, 0)
        .Weight     = REDBOX_LINE_WEIGHT
        .DashStyle  = msoLineSolid
    End With

    ' Füllung
    If filled Then
        With shp.Fill
            .Visible        = msoTrue
            .Solid
            .ForeColor.RGB  = RGB(204, 0, 0)
            .Transparency   = FILL_TRANSPARENCY
        End With
    Else
        shp.Fill.Visible = msoFalse
    End If

    ' Kein Text, kein Textrahmen-Innenabstand
    On Error Resume Next
    shp.TextFrame.TextRange.Text = ""
    On Error GoTo ErrHandler

    Set CreateRedBoxShape = shp
    Exit Function

ErrHandler:
    Set CreateRedBoxShape = Nothing
End Function


' =============================================================================
' HILFS-FUNKTIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Berechnet Bounding-Box aller selektierten Shapes + Padding.
' Parameter:  padding          - Abstand um die Selektion (pt)
'             left, top, w, h  - Ausgabe-Werte
' -----------------------------------------------------------------------
Public Sub GetSelectionBounds(padding As Single, _
                               ByRef left As Single, ByRef top As Single, _
                               ByRef w As Single, ByRef h As Single)

    On Error GoTo Fallback

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    Dim minL As Single, minT As Single
    Dim maxR As Single, maxB As Single

    minL = sr(1).Left
    minT = sr(1).Top
    maxR = sr(1).Left + sr(1).Width
    maxB = sr(1).Top  + sr(1).Height

    Dim i As Long
    For i = 2 To sr.Count
        If sr(i).Left < minL Then minL = sr(i).Left
        If sr(i).Top  < minT Then minT = sr(i).Top
        If sr(i).Left + sr(i).Width  > maxR Then maxR = sr(i).Left + sr(i).Width
        If sr(i).Top  + sr(i).Height > maxB Then maxB = sr(i).Top  + sr(i).Height
    Next i

    left = minL - padding
    top  = minT - padding
    w    = (maxR - minL) + 2 * padding
    h    = (maxB - minT) + 2 * padding
    Exit Sub

Fallback:
    CenterBoxOnSlide left, top, w, h
End Sub


' -----------------------------------------------------------------------
' Zweck:      Berechnet zentrierte Position für Default-Box.
' -----------------------------------------------------------------------
Private Sub CenterBoxOnSlide(ByRef left As Single, ByRef top As Single, _
                              ByRef w As Single, ByRef h As Single)
    On Error Resume Next
    Dim slideW As Single
    Dim slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    On Error GoTo 0

    If slideW = 0 Then slideW = 720   ' PPT-Standard 10 Zoll = 720 pt
    If slideH = 0 Then slideH = 540   ' PPT-Standard 7.5 Zoll = 540 pt

    w    = DEFAULT_W
    h    = DEFAULT_H
    left = (slideW - w) / 2
    top  = (slideH - h) / 2
End Sub


' -----------------------------------------------------------------------
' Zweck:      Gibt True zurück wenn Shapes selektiert sind.
' -----------------------------------------------------------------------
Private Function HasShapesSelected() As Boolean
    On Error Resume Next
    HasShapesSelected = (ActiveWindow.Selection.Type = ppSelectionShapes) And _
                        (ActiveWindow.Selection.ShapeRange.Count > 0)
    On Error GoTo 0
End Function


' -----------------------------------------------------------------------
' Zweck:      Zählt Red Boxes auf einer Folie.
' -----------------------------------------------------------------------
Private Function CountRedBoxesOnSlide(sld As Slide) As Long
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 1 To sld.Shapes.Count
        On Error Resume Next
        Dim tagVal As String
        tagVal = sld.Shapes(i).Tags(REDBOX_TAG_KEY)
        On Error GoTo 0
        If tagVal = REDBOX_TAG_VALUE Then count = count + 1
    Next i
    CountRedBoxesOnSlide = count
End Function


' -----------------------------------------------------------------------
' Zweck:      Löscht alle Red Boxes von einer Folie.
' Rückgabe:   Anzahl gelöschter Shapes
' -----------------------------------------------------------------------
Private Function DeleteRedBoxesFromSlide(sld As Slide) As Long
    Dim removed As Long
    removed = 0
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        On Error Resume Next
        Dim tagVal As String
        tagVal = sld.Shapes(i).Tags(REDBOX_TAG_KEY)
        On Error GoTo 0
        If tagVal = REDBOX_TAG_VALUE Then
            sld.Shapes(i).Delete
            removed = removed + 1
        End If
    Next i
    DeleteRedBoxesFromSlide = removed
End Function
