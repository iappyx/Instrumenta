Attribute VB_Name = "modFormatPainterPlus"
Option Explicit

' =============================================================================
' Modul:  modFormatPainterPlus
' Zweck:  Format Painter Plus – kopiert Formatierung von einem Quell-Shape und
'         wendet selektierte Eigenschaften auf mehrere Ziel-Shapes an.
'
' Workflow:
'   1. Genau ein Shape selektieren (Quelle) → Button klicken
'   2. Form öffnet sich mit Übersicht der gecapturten Werte und Checkboxen
'   3. Zu übertragende Eigenschaften per Checkbox wählen
'   4. Ziel-Shapes selektieren (kann auf anderer Folie sein)
'   5. "Anwenden" klicken
'
' Plattform:     Windows und Mac
' Hinweis:       UndoRecord steht in PowerPoint VBA nicht zur Verfügung.
'                Jede Shape-Änderung erzeugt einen eigenen Undo-Eintrag.
' =============================================================================

' --- Gecapturte Quell-Formatierung (modulweit, von Form lesbar) --------------

Public Type FormatSnapshot
    ' Quelle
    SourceName      As String

    ' Füllung
    FillVisible     As Boolean
    FillType        As Long             ' MsoFillType-Wert
    FillColor       As Long             ' RGB
    FillTransp      As Single           ' 0.0–1.0

    ' Linie
    LineVisible     As Boolean
    LineColor       As Long             ' RGB
    LineWeight      As Single           ' pt
    LineDash        As Long             ' MsoLineDashStyle

    ' Schrift (aus erstem Lauf des ersten Paragraphen)
    FontName        As String
    FontSize        As Single           ' pt
    FontBold        As Boolean
    FontItalic      As Boolean
    FontUnderline   As Boolean
    FontColor       As Long             ' RGB

    ' Text-Ausrichtung
    TextAlignH      As Long             ' PpParagraphAlignment
    TextAlignV      As Long             ' PpTextVerticalAlignment

    ' Maße
    ShapeWidth      As Single           ' pt
    ShapeHeight     As Single           ' pt
End Type

Public g_Snapshot As FormatSnapshot
Public g_SnapshotValid As Boolean

' --- Apply-Optionen (hier deklariert damit alle Sub/Function-Signaturen stimmen)

Public Type ApplyOptions
    FillColor       As Boolean
    LineColor       As Boolean
    LineWeight      As Boolean
    LineDash        As Boolean
    FontName        As Boolean
    FontSize        As Boolean
    FontBold        As Boolean
    FontItalic      As Boolean
    FontUnderline   As Boolean
    FontColor       As Boolean
    TextAlignH      As Boolean
    TextAlignV      As Boolean
    ShapeWidth      As Boolean
    ShapeHeight     As Boolean
End Type


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – liest Formatierung vom selektierten Quell-Shape
'             und öffnet die Form.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub ShowFormatPainterPlus(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Format Painter+"

    On Error GoTo ErrHandler

    ' Genau ein Shape muss selektiert sein
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Bitte genau ein Shape als Quelle selektieren.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    If sel.ShapeRange.Count <> 1 Then
        MsgBox "Bitte genau ein Shape als Quelle selektieren (" & _
               sel.ShapeRange.Count & " selektiert).", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' Formatierung capturen
    CaptureFormat sel.ShapeRange(1)

    ' Form öffnen
    frmFormatPainterPlus.InitForm
    frmFormatPainterPlus.Show vbModeless

    Exit Sub
ErrHandler:
    MsgBox "Fehler in ShowFormatPainterPlus: " & Err.Description, _
           vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Liest alle unterstützten Formateigenschaften vom Shape.
' Parameter:  shp - Quell-Shape
' -----------------------------------------------------------------------
Public Sub CaptureFormat(shp As Shape)

    g_SnapshotValid = False

    With g_Snapshot
        .SourceName = shp.Name

        ' --- Füllung
        On Error Resume Next
        .FillVisible = (shp.Fill.Visible = msoTrue)
        .FillType = shp.Fill.Type
        .FillColor = shp.Fill.ForeColor.RGB
        .FillTransp = shp.Fill.Transparency
        On Error GoTo 0

        ' --- Linie
        On Error Resume Next
        .LineVisible = (shp.Line.Visible = msoTrue)
        .LineColor = shp.Line.ForeColor.RGB
        .LineWeight = shp.Line.Weight
        .LineDash = shp.Line.DashStyle
        On Error GoTo 0

        ' --- Schrift (aus erstem Lauf des ersten Paragraphen)
        On Error Resume Next
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Dim tr As TextRange
                Set tr = shp.TextFrame.TextRange
                If tr.Paragraphs.Count > 0 Then
                    Dim firstPara As TextRange
                    Set firstPara = tr.Paragraphs(1)
                    Dim firstRun As TextRange
                    If firstPara.Runs.Count > 0 Then
                        Set firstRun = firstPara.Runs(1)
                        .FontName = firstRun.Font.Name
                        .FontSize = firstRun.Font.Size
                        .FontBold = (firstRun.Font.Bold = msoTrue)
                        .FontItalic = (firstRun.Font.Italic = msoTrue)
                        .FontUnderline = (firstRun.Font.Underline = msoTrue)
                        .FontColor = firstRun.Font.Color.RGB
                    End If
                    .TextAlignH = firstPara.ParagraphFormat.Alignment
                End If
                .TextAlignV = shp.TextFrame.VerticalAnchor
            End If
        End If
        On Error GoTo 0

        ' --- Maße
        On Error Resume Next
        .ShapeWidth = shp.Width
        .ShapeHeight = shp.Height
        On Error GoTo 0
    End With

    g_SnapshotValid = True
End Sub


' -----------------------------------------------------------------------
' Zweck:      Wendet die gecapturten Eigenschaften auf alle selektierten
'             Shapes an. Wird von frmFormatPainterPlus aufgerufen.
' Parameter:  opts - Steuert welche Eigenschaften angewendet werden
' -----------------------------------------------------------------------
Public Sub ApplyFormatToSelection(opts As ApplyOptions)

    Const DLG_TITLE As String = "Infront Toolkit – Format Painter+"

    On Error GoTo ErrHandler

    If Not g_SnapshotValid Then
        MsgBox "Keine gecapturte Formatierung vorhanden.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Bitte Ziel-Shapes selektieren.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    Dim adjusted As Long
    Dim skipped As Long
    Dim i As Long

    For i = 1 To sr.Count
        Dim shp As Shape
        Set shp = sr(i)

        If ApplyToShape(shp, opts) Then
            adjusted = adjusted + 1
        Else
            skipped = skipped + 1
        End If
    Next i

    Dim msg As String
    msg = adjusted & " Shape" & IIf(adjusted = 1, "", "s") & " angepasst."
    If skipped > 0 Then
        msg = msg & vbCrLf & skipped & " übersprungen (nicht unterstützte Eigenschaften)."
    End If
    MsgBox msg, vbInformation, DLG_TITLE

    Exit Sub
ErrHandler:
    MsgBox "Fehler beim Anwenden: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Wendet Optionen auf ein einzelnes Shape an.
' Parameter:  shp  - Ziel-Shape
'             opts - anzuwendende Eigenschaften
' Rückgabe:   True wenn mindestens eine Eigenschaft angewendet wurde
' -----------------------------------------------------------------------
Private Function ApplyToShape(shp As Shape, opts As ApplyOptions) As Boolean

    Dim applied As Boolean
    applied = False

    ' --- Füllung
    If opts.FillColor Then
        On Error Resume Next
        If g_Snapshot.FillVisible Then
            shp.Fill.Visible = msoTrue
            If g_Snapshot.FillType = msoFillSolid Then
                shp.Fill.Solid
                shp.Fill.ForeColor.RGB = g_Snapshot.FillColor
                shp.Fill.Transparency = g_Snapshot.FillTransp
            End If
        Else
            shp.Fill.Visible = msoFalse
        End If
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Linienfarbe
    If opts.LineColor Then
        On Error Resume Next
        If g_Snapshot.LineVisible Then
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = g_Snapshot.LineColor
        Else
            shp.Line.Visible = msoFalse
        End If
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Linienstärke
    If opts.LineWeight Then
        On Error Resume Next
        If g_Snapshot.LineVisible Then
            shp.Line.Weight = g_Snapshot.LineWeight
        End If
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Linienstil
    If opts.LineDash Then
        On Error Resume Next
        If g_Snapshot.LineVisible Then
            shp.Line.DashStyle = g_Snapshot.LineDash
        End If
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Schriftart
    If opts.FontName Then
        On Error Resume Next
        If shp.HasTextFrame Then
            If g_Snapshot.FontName <> "" Then
                shp.TextFrame.TextRange.Font.Name = g_Snapshot.FontName
                If Err.Number = 0 Then applied = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Schriftgröße
    If opts.FontSize Then
        On Error Resume Next
        If shp.HasTextFrame Then
            If g_Snapshot.FontSize > 0 Then
                shp.TextFrame.TextRange.Font.Size = g_Snapshot.FontSize
                If Err.Number = 0 Then applied = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Fett
    If opts.FontBold Then
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Font.Bold = _
                IIf(g_Snapshot.FontBold, msoTrue, msoFalse)
            If Err.Number = 0 Then applied = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Kursiv
    If opts.FontItalic Then
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Font.Italic = _
                IIf(g_Snapshot.FontItalic, msoTrue, msoFalse)
            If Err.Number = 0 Then applied = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Unterstrichen
    If opts.FontUnderline Then
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Font.Underline = _
                IIf(g_Snapshot.FontUnderline, msoTrue, msoFalse)
            If Err.Number = 0 Then applied = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Schriftfarbe
    If opts.FontColor Then
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.Font.Color.RGB = g_Snapshot.FontColor
            If Err.Number = 0 Then applied = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Horizontale Textausrichtung
    If opts.TextAlignH Then
        On Error Resume Next
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Dim p As Long
                For p = 1 To shp.TextFrame.TextRange.Paragraphs.Count
                    shp.TextFrame.TextRange.Paragraphs(p).ParagraphFormat.Alignment _
                        = g_Snapshot.TextAlignH
                Next p
                If Err.Number = 0 Then applied = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Vertikale Textausrichtung
    If opts.TextAlignV Then
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.VerticalAnchor = g_Snapshot.TextAlignV
            If Err.Number = 0 Then applied = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Breite
    If opts.ShapeWidth Then
        On Error Resume Next
        shp.Width = g_Snapshot.ShapeWidth
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Höhe
    If opts.ShapeHeight Then
        On Error Resume Next
        shp.Height = g_Snapshot.ShapeHeight
        If Err.Number = 0 Then applied = True
        Err.Clear
        On Error GoTo 0
    End If

    ApplyToShape = applied
End Function


' (ApplyOptions-Typ ist oben im Modul deklariert)
