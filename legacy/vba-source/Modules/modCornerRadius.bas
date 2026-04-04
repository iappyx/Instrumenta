Attribute VB_Name = "modCornerRadius"
Option Explicit

' =============================================================================
' Modul:  modCornerRadius
' Zweck:  Setzt den Eckenradius ausgewählter Shapes anhand eines Pixelwerts.
'         Unterstützt Shapes mit Eckenradius-Justierung (z.B. abgerundete
'         Rechtecke). Nicht unterstützte Shapes werden übersprungen.
' Plattform: Windows und Mac (kein Windows-API-Aufruf)
' Undo:   PowerPoint erstellt automatisch pro Shape-Änderung einen Undo-Eintrag.
'         Mehrere Shapes = mehrere Undo-Schritte (PPT VBA hat kein UndoRecord).
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – fragt Eckenradius in Pixel ab und setzt
'             ihn auf alle ausgewählten Shapes, die Justierungen unterstützen.
' Parameter:  control - IRibbonControl (wird vom Ribbon übergeben)
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub SetCornerRadiusPx(control As IRibbonControl)

    Const PROC_NAME As String = "SetCornerRadiusPx"
    Const DLG_TITLE As String = "Infront Toolkit – Eckenradius"

    On Error GoTo ErrHandler

    ' -- 1. Selektion prüfen
    Dim sel As Selection
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    On Error GoTo ErrHandler

    If sel Is Nothing Then
        MsgBox "Keine aktive Folie gefunden.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    If sel.Type <> ppSelectionShapes Then
        MsgBox "Bitte zuerst mindestens eine Shape auswählen.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' -- 2. Pixelwert abfragen
    Dim inputVal As String
    inputVal = InputBox("Eckenradius in Pixel (z.B. 8):", DLG_TITLE, "8")

    If inputVal = "" Then Exit Sub   ' Abbruch durch Nutzer

    If Not IsNumeric(inputVal) Then
        MsgBox "Ungültige Eingabe. Bitte eine positive Zahl eingeben.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim radiusPx As Double
    radiusPx = CDbl(inputVal)

    If radiusPx < 0 Then
        MsgBox "Der Eckenradius muss 0 oder größer sein.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' -- 3. Pixel → Punkte (96 DPI: 1 px = 0,75 pt)
    Dim radiusPt As Double
    radiusPt = radiusPx * 0.75

    ' -- 4. ShapeRange bestimmen (direkte Auswahl oder Child-Shapes in Gruppe)
    Dim sr As ShapeRange
    If sel.HasChildShapeRange Then
        Set sr = sel.ChildShapeRange
    Else
        Set sr = sel.ShapeRange
    End If

    ' -- 5. Auf alle Shapes anwenden
    Dim shp As Shape
    Dim adjustedCount As Long
    Dim skippedCount As Long
    adjustedCount = 0
    skippedCount = 0

    Dim i As Long
    For i = 1 To sr.Count
        Set shp = sr(i)
        If ApplyCornerRadius(shp, radiusPt) Then
            adjustedCount = adjustedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next i

    ' -- 6. Ergebnismeldung
    Dim msg As String
    If adjustedCount = 0 Then
        msg = "Kein Shape konnte angepasst werden." & vbCrLf & _
              "Nur Shapes mit Eckenradius-Justierung werden unterstützt" & vbCrLf & _
              "(z.B. abgerundete Rechtecke)."
        MsgBox msg, vbInformation, DLG_TITLE
    ElseIf skippedCount = 0 Then
        msg = adjustedCount & " Shape(s) auf " & radiusPx & " px Eckenradius gesetzt."
        MsgBox msg, vbInformation, DLG_TITLE
    Else
        msg = adjustedCount & " Shape(s) angepasst." & vbCrLf & _
              skippedCount & " Shape(s) übersprungen (kein Eckenradius-Support)."
        MsgBox msg, vbInformation, DLG_TITLE
    End If

    Exit Sub

ErrHandler:
    MsgBox "Fehler in " & PROC_NAME & ": " & Err.Description & _
           " (Nr. " & Err.Number & ")", vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Berechnet und setzt den Eckenradius-Adjustment eines Shapes.
'             Shapes ohne Adjustments oder ohne Eckenradius-Support werden
'             still übersprungen (kein Fehler nach außen).
' Parameter:  shp      - das zu ändernde Shape
'             radiusPt - gewünschter Radius in Punkten
' Rückgabe:   True wenn der Radius erfolgreich gesetzt wurde, sonst False
' -----------------------------------------------------------------------
Private Function ApplyCornerRadius(shp As Shape, radiusPt As Double) As Boolean

    On Error GoTo NotSupported

    ' Adjustments verfügbar?
    Dim adjCount As Long
    adjCount = shp.Adjustments.Count
    If adjCount < 1 Then
        ApplyCornerRadius = False
        Exit Function
    End If

    ' Min(Width, Height) bestimmen
    Dim minDim As Double
    If shp.Width < shp.Height Then
        minDim = shp.Width
    Else
        minDim = shp.Height
    End If

    If minDim <= 0 Then
        ApplyCornerRadius = False
        Exit Function
    End If

    ' Normierter Adjustment-Wert = radiusPt / (minDim / 2), gedeckelt auf 0.5
    ' Formel konsistent mit ModuleObjectsRoundedCorners.bas
    Dim adjVal As Double
    adjVal = radiusPt / (minDim / 2)
    If adjVal > 0.5 Then adjVal = 0.5
    If adjVal < 0 Then adjVal = 0

    ' Setzen – bei nicht unterstützten Shapes springt On Error zu NotSupported
    shp.Adjustments(1) = adjVal

    ApplyCornerRadius = True
    On Error GoTo 0
    Exit Function

NotSupported:
    ' Shape unterstützt keinen Eckenradius (z.B. Linie, Tabelle, Gruppe, etc.)
    ApplyCornerRadius = False
    On Error GoTo 0
End Function
