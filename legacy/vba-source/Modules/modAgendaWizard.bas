Attribute VB_Name = "modAgendaWizard"
Option Explicit

' =============================================================================
' Modul:  modAgendaWizard
' Zweck:  Agenda Wizard – erstellt Übersichts- und optionale Fortschrittsfolien
'         aus einer benutzerdefinierten Themen-Liste.
'
' Konzept:
'   - Eine Master-Übersichtsfolie (alle Punkte aktiv/gleiche Farbe)
'   - Optional: Fortschrittsfolien vor jeder Sektion
'     (aktueller Punkt hervorgehoben, bereits behandelte normal,
'      kommende Punkte gedimmt)
'   - Agenda-Folien werden mit Tag "InfrontAgenda" markiert →
'     idempotent: Neugenerierung löscht vorherige automatisch
'   - Alle Shapes per VBA positioniert (keine Layout-Abhängigkeit)
'
' Plattform:  Windows und Mac
' =============================================================================

' --- Öffentlicher Konfigurations-Typ (von frmAgendaWizard befüllt) -----------

Public Type AgendaConfig
    Title           As String       ' Folientitel
    Items()         As String       ' Agendapunkte (1-basiert)
    ItemCount       As Long
    ActiveColor     As Long         ' RGB – aktiver/aktueller Punkt
    InactiveColor   As Long         ' RGB – inaktive/zukünftige Punkte
    DoneColor       As Long         ' RGB – bereits behandelte Punkte
    TitleFontSize   As Single       ' pt, 0 = auto (24)
    ItemFontSize    As Single       ' pt, 0 = auto (16)
    ' InsertionMode: 0 = Nur Übersichtsfolie
    '                1 = Übersicht + Fortschrittsfolie vor jeder Sektion
    InsertionMode   As Long
    InsertAfterSlide As Long        ' 0 = Anfang der Präsentation
End Type

' Tag-Konstante für Wiedererkennung
Private Const AGENDA_TAG_KEY   As String = "InfrontAgenda"
Private Const AGENDA_TAG_VALUE As String = "1"


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – öffnet den Agenda Wizard.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub ShowAgendaWizard(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Agenda Wizard"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    frmAgendaWizard.InitForm
    frmAgendaWizard.Show vbModeless

    Exit Sub
ErrHandler:
    MsgBox "Fehler in ShowAgendaWizard: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Löscht vorhandene Agenda-Folien und generiert neue.
' Parameter:  cfg - AgendaConfig
' -----------------------------------------------------------------------
Public Sub GenerateAgenda(cfg As AgendaConfig)

    Const DLG_TITLE As String = "Infront Toolkit – Agenda Wizard"

    On Error GoTo ErrHandler

    If cfg.ItemCount = 0 Then
        MsgBox "Bitte mindestens einen Agendapunkt eingeben.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    If cfg.Title = "" Then cfg.Title = "Agenda"
    If cfg.TitleFontSize <= 0 Then cfg.TitleFontSize = 24
    If cfg.ItemFontSize <= 0 Then cfg.ItemFontSize = 16
    If cfg.ActiveColor = 0 Then cfg.ActiveColor = RGB(0, 51, 102)
    If cfg.InactiveColor = 0 Then cfg.InactiveColor = RGB(180, 180, 180)
    If cfg.DoneColor = 0 Then cfg.DoneColor = RGB(100, 100, 100)

    ' Bestehende Agenda-Folien löschen
    DeleteExistingAgendaSlides

    ' Einfügeposition nach Löschung neu berechnen
    Dim insertPos As Long
    insertPos = cfg.InsertAfterSlide + 1  ' 1-basiert

    If insertPos < 1 Then insertPos = 1
    If insertPos > ActivePresentation.Slides.Count + 1 Then
        insertPos = ActivePresentation.Slides.Count + 1
    End If

    ' --- Modus 0: Nur Master-Übersichtsfolie
    If cfg.InsertionMode = 0 Then
        InsertAgendaSlide insertPos, cfg, -1   ' -1 = alle aktiv
        insertPos = insertPos + 1
    End If

    ' --- Modus 1: Übersicht + Fortschrittsfolien
    If cfg.InsertionMode = 1 Then
        ' Master-Übersicht (alle aktiv)
        InsertAgendaSlide insertPos, cfg, -1
        insertPos = insertPos + 1

        ' Fortschrittsfolie vor Sektion i: Punkt i aktiv, i-1..1 done, i+1..n inaktiv
        Dim i As Long
        For i = 1 To cfg.ItemCount
            InsertAgendaSlide insertPos, cfg, i
            insertPos = insertPos + 1
        Next i
    End If

    Dim total As Long
    If cfg.InsertionMode = 0 Then
        total = 1
    Else
        total = 1 + cfg.ItemCount
    End If

    MsgBox total & " Agenda-Folie" & IIf(total = 1, "", "n") & " eingefügt.", _
           vbInformation, DLG_TITLE

    Exit Sub
ErrHandler:
    MsgBox "Fehler in GenerateAgenda: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Löscht alle mit InfrontAgenda=1 markierten Folien.
' Rückgabe:   Anzahl gelöschter Folien
' -----------------------------------------------------------------------
Public Function DeleteExistingAgendaSlides() As Long

    Dim deleted As Long
    deleted = 0

    ' Rückwärts löschen um Index-Verschiebung zu vermeiden
    Dim i As Long
    For i = ActivePresentation.Slides.Count To 1 Step -1
        On Error Resume Next
        Dim tagVal As String
        tagVal = ActivePresentation.Slides(i).Tags(AGENDA_TAG_KEY)
        On Error GoTo 0

        If tagVal = AGENDA_TAG_VALUE Then
            ActivePresentation.Slides(i).Delete
            deleted = deleted + 1
        End If
    Next i

    DeleteExistingAgendaSlides = deleted
End Function


' =============================================================================
' FOLIE GENERIEREN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Fügt eine einzelne Agenda-Folie an gegebener Position ein.
' Parameter:  insertPos  - Position (1-basiert)
'             cfg        - Konfiguration
'             activeIdx  - Index des hervorgehobenen Punktes (1-basiert)
'                          -1 = alle Punkte aktiv (Master-Übersicht)
' -----------------------------------------------------------------------
Private Sub InsertAgendaSlide(insertPos As Long, cfg As AgendaConfig, _
                               activeIdx As Long)

    On Error GoTo ErrHandler

    ' Neue leere Folie einfügen
    Dim sld As Slide
    Set sld = ActivePresentation.Slides.Add(insertPos, ppLayoutBlank)

    ' Tag setzen für späteres Wiederfinden
    sld.Tags.Add AGENDA_TAG_KEY, AGENDA_TAG_VALUE

    ' Foliengröße ermitteln
    Dim slideW As Single
    Dim slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    ' --- Hintergrund (weiß/Standard – kein Override nötig)

    ' --- Titel einfügen
    Dim margin As Single
    margin = 36  ' pt (ca. 1.27 cm)

    Dim titleBox As Shape
    Set titleBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        margin, margin, slideW - 2 * margin, cfg.TitleFontSize + 12)

    With titleBox.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText
        With .TextRange
            .Text = cfg.Title
            With .Font
                .Name = "Calibri"
                .Size = cfg.TitleFontSize
                .Bold = msoTrue
                .Color.RGB = cfg.ActiveColor
            End With
        End With
    End With

    ' --- Divider-Linie unter Titel
    Dim lineY As Single
    lineY = titleBox.Top + titleBox.Height + 6

    Dim divLine As Shape
    Set divLine = sld.Shapes.AddLine(margin, lineY, slideW - margin, lineY)
    divLine.Line.ForeColor.RGB = cfg.ActiveColor
    divLine.Line.Weight = 1.5

    ' --- Agendapunkte
    Dim itemAreaTop As Single
    itemAreaTop = lineY + 12

    Dim itemAreaH As Single
    itemAreaH = slideH - itemAreaTop - margin

    ' Zeilenhöhe berechnen
    Dim rowH As Single
    rowH = itemAreaH / cfg.ItemCount
    If rowH > cfg.ItemFontSize + 18 Then rowH = cfg.ItemFontSize + 18
    If rowH < cfg.ItemFontSize + 4 Then rowH = cfg.ItemFontSize + 4

    Dim i As Long
    For i = 1 To cfg.ItemCount

        ' Farbe je nach Status
        Dim itemColor As Long
        If activeIdx = -1 Then
            ' Master-Übersicht: alle gleich aktiv
            itemColor = cfg.ActiveColor
        ElseIf i = activeIdx Then
            itemColor = cfg.ActiveColor
        ElseIf i < activeIdx Then
            itemColor = cfg.DoneColor
        Else
            itemColor = cfg.InactiveColor
        End If

        ' Nummernbox (links)
        Dim numW As Single
        numW = cfg.ItemFontSize * 1.8

        Dim numBox As Shape
        Set numBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            margin, _
            itemAreaTop + (i - 1) * rowH, _
            numW, rowH)

        With numBox.TextFrame
            .AutoSize = ppAutoSizeNone
            With .TextRange
                .Text = CStr(i) & "."
                With .ParagraphFormat
                    .Alignment = ppAlignRight
                End With
                With .Font
                    .Name = "Calibri"
                    .Size = cfg.ItemFontSize
                    .Bold = IIf(i = activeIdx And activeIdx <> -1, msoTrue, msoFalse)
                    .Color.RGB = itemColor
                End With
            End With
        End With

        ' Textbox für Agendapunkt
        Dim itemBox As Shape
        Set itemBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            margin + numW + 6, _
            itemAreaTop + (i - 1) * rowH, _
            slideW - 2 * margin - numW - 6, rowH)

        With itemBox.TextFrame
            .AutoSize = ppAutoSizeNone
            With .TextRange
                .Text = cfg.Items(i)
                With .ParagraphFormat
                    .Alignment = ppAlignLeft
                End With
                With .Font
                    .Name = "Calibri"
                    .Size = cfg.ItemFontSize
                    .Bold = IIf(i = activeIdx And activeIdx <> -1, msoTrue, msoFalse)
                    .Color.RGB = itemColor
                End With
            End With
        End With

    Next i

    Exit Sub
ErrHandler:
    ' Fehlgeschlagene Folien werden übersprungen – keine Unterbrechung
    On Error Resume Next
    If Not sld Is Nothing Then
        If sld.SlideIndex > 0 Then sld.Delete
    End If
End Sub


' =============================================================================
' HILFS-FUNKTIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Parst mehrzeiligen Text in String-Array (eine Zeile = ein Punkt).
'             Leere Zeilen werden übersprungen.
' Parameter:  raw       - Eingabe-String (vbCrLf / vbLf getrennt)
'             items()   - Ausgabe-Array (1-basiert)
'             itemCount - Anzahl gültiger Einträge
' -----------------------------------------------------------------------
Public Sub ParseItemList(raw As String, ByRef items() As String, _
                         ByRef itemCount As Long)

    itemCount = 0

    ' Normalisieren: \r\n → \n
    Dim normalized As String
    normalized = Replace(raw, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim lines() As String
    lines = Split(normalized, vbLf)

    ReDim items(1 To UBound(lines) + 1)

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        If Len(line) > 0 Then
            itemCount = itemCount + 1
            items(itemCount) = line
        End If
    Next i
End Sub


' -----------------------------------------------------------------------
' Zweck:      Gibt Anzahl vorhandener Agenda-Folien zurück.
' -----------------------------------------------------------------------
Public Function CountAgendaSlides() As Long
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 1 To ActivePresentation.Slides.Count
        On Error Resume Next
        Dim tagVal As String
        tagVal = ActivePresentation.Slides(i).Tags(AGENDA_TAG_KEY)
        On Error GoTo 0
        If tagVal = AGENDA_TAG_VALUE Then count = count + 1
    Next i
    CountAgendaSlides = count
End Function
