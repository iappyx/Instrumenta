Attribute VB_Name = "modUserStamp"
Option Explicit

' =============================================================================
' Modul:  modUserStamp
' Zweck:  User-Name-Stempel – setzt einen kleinen Review-Stempel mit Name
'         und Datum/Uhrzeit auf eine oder mehrere Folien.
'
' Stempel-Format:  "[Name]  |  DD.MM.YYYY  HH:MM"
' Position:        Unten rechts auf der Folie
' Erkennung:       Tag InfrontUserStamp=1
' Idempotenz:      Pro Folie maximal ein Stempel – existierender wird ersetzt
'
' Plattform:       Windows und Mac
' =============================================================================

Private Const STAMP_TAG_KEY   As String = "InfrontUserStamp"
Private Const STAMP_TAG_VALUE As String = "1"
Private Const STAMP_WIDTH     As Single = 220   ' pt
Private Const STAMP_HEIGHT    As Single = 14    ' pt
Private Const STAMP_MARGIN    As Single = 8     ' pt vom Folienrand
Private Const STAMP_FONT_SIZE As Single = 7.5
Private Const STAMP_FONT_NAME As String = "Calibri"


' =============================================================================
' RIBBON-CALLBACKS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – fügt Stempel ein.
'             Scope-Auswahl: Aktuelle Folie / Alle / Selektierte.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub InsertUserStamp(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – User-Stempel"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' Stempeltext aufbauen
    Dim stampText As String
    stampText = BuildStampText()
    If stampText = "" Then Exit Sub  ' Abbruch durch Nutzer

    ' Scope wählen
    Dim scope As Long
    scope = AskScope(DLG_TITLE)
    If scope = -1 Then Exit Sub  ' Abbruch

    ' Stempel einfügen
    Dim slides As SlideRange
    Set slides = GetScopeSlides(scope)
    If slides Is Nothing Then Exit Sub

    Dim inserted As Long
    inserted = 0

    Dim sld As Slide
    For Each sld In slides
        AddStampToSlide sld, stampText
        inserted = inserted + 1
    Next sld

    MsgBox "Stempel auf " & inserted & " Folie" & IIf(inserted = 1, "", "n") & " gesetzt.", _
           vbInformation, DLG_TITLE

    Exit Sub
ErrHandler:
    MsgBox "Fehler in InsertUserStamp: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – entfernt alle User-Stempel aus allen Folien.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub RemoveUserStamps(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – User-Stempel"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim count As Long
    count = CountStamps()

    If count = 0 Then
        MsgBox "Keine User-Stempel vorhanden.", vbInformation, DLG_TITLE
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox(count & " Stempel entfernen?", _
                    vbQuestion + vbOKCancel, DLG_TITLE)
    If answer <> vbOK Then Exit Sub

    Dim removed As Long
    removed = DeleteAllStamps()

    MsgBox removed & " Stempel entfernt.", vbInformation, DLG_TITLE

    Exit Sub
ErrHandler:
    MsgBox "Fehler in RemoveUserStamps: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' =============================================================================
' STEMPEL-OPERATIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Fügt Stempel auf einer Folie ein. Existierender Stempel
'             auf dieser Folie wird vorher entfernt (Idempotenz).
' Parameter:  sld       - Zielfolie
'             stampText - anzuzeigender Text
' -----------------------------------------------------------------------
Public Sub AddStampToSlide(sld As Slide, stampText As String)

    On Error Resume Next

    ' Vorhandenen Stempel auf dieser Folie entfernen
    RemoveStampFromSlide sld

    ' Folienmaße
    Dim slideW As Single
    Dim slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    ' Position: unten rechts
    Dim left As Single
    Dim top As Single
    left = slideW - STAMP_WIDTH - STAMP_MARGIN
    top  = slideH - STAMP_HEIGHT - STAMP_MARGIN

    ' Textbox hinzufügen
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        left, top, STAMP_WIDTH, STAMP_HEIGHT)

    If shp Is Nothing Then Exit Sub

    ' Text und Format
    With shp
        .Tags.Add STAMP_TAG_KEY, STAMP_TAG_VALUE

        With .TextFrame
            .AutoSize = ppAutoSizeNone
            .WordWrap = msoFalse
            .MarginLeft   = 0
            .MarginRight  = 0
            .MarginTop    = 0
            .MarginBottom = 0

            With .TextRange
                .Text = stampText
                With .ParagraphFormat
                    .Alignment = ppAlignRight
                End With
                With .Font
                    .Name  = STAMP_FONT_NAME
                    .Size  = STAMP_FONT_SIZE
                    .Bold  = msoFalse
                    .Color.RGB = RGB(150, 150, 150)
                End With
            End With
        End With

        ' Kein Rahmen, kein Hintergrund
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Entfernt den Stempel von einer einzelnen Folie.
' Parameter:  sld - Folie
' -----------------------------------------------------------------------
Private Sub RemoveStampFromSlide(sld As Slide)

    On Error Resume Next

    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        Dim tagVal As String
        tagVal = sld.Shapes(i).Tags(STAMP_TAG_KEY)
        If tagVal = STAMP_TAG_VALUE Then
            sld.Shapes(i).Delete
        End If
    Next i

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Entfernt alle Stempel aus der gesamten Präsentation.
' Rückgabe:   Anzahl entfernter Stempel
' -----------------------------------------------------------------------
Private Function DeleteAllStamps() As Long

    Dim removed As Long
    removed = 0

    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim i As Long
        For i = sld.Shapes.Count To 1 Step -1
            On Error Resume Next
            Dim tagVal As String
            tagVal = sld.Shapes(i).Tags(STAMP_TAG_KEY)
            On Error GoTo 0
            If tagVal = STAMP_TAG_VALUE Then
                sld.Shapes(i).Delete
                removed = removed + 1
            End If
        Next i
    Next sld

    DeleteAllStamps = removed
End Function


' -----------------------------------------------------------------------
' Zweck:      Zählt vorhandene Stempel in der Präsentation.
' -----------------------------------------------------------------------
Private Function CountStamps() As Long

    Dim count As Long
    count = 0

    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim i As Long
        For i = 1 To sld.Shapes.Count
            On Error Resume Next
            Dim tagVal As String
            tagVal = sld.Shapes(i).Tags(STAMP_TAG_KEY)
            On Error GoTo 0
            If tagVal = STAMP_TAG_VALUE Then count = count + 1
        Next i
    Next sld

    CountStamps = count
End Function


' =============================================================================
' HILFS-FUNKTIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Baut den Stempel-Text aus Username und aktuellem Datum/Zeit.
'             Fragt per InputBox nach dem Namen wenn Application.UserName leer.
' Rückgabe:   Stempel-String oder "" bei Abbruch
' -----------------------------------------------------------------------
Public Function BuildStampText() As String

    Dim userName As String
    On Error Resume Next
    userName = Trim(Application.UserName)
    On Error GoTo 0

    If userName = "" Then
        userName = InputBox("Bitte Namen eingeben:", _
                            "Infront Toolkit – User-Stempel", "")
        userName = Trim(userName)
        If userName = "" Then
            BuildStampText = ""
            Exit Function
        End If
    End If

    BuildStampText = userName & "  |  " & Format(Now, "DD.MM.YYYY  HH:MM")
End Function


' -----------------------------------------------------------------------
' Zweck:      Fragt den Scope per MsgBox ab.
' Rückgabe:   0 = Alle, 1 = Selektierte, 2 = Aktuelle Folie, -1 = Abbruch
' -----------------------------------------------------------------------
Private Function AskScope(dlgTitle As String) As Long

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Wo soll der Stempel eingefügt werden?" & vbCrLf & vbCrLf & _
                    "Ja       = Aktuelle Folie" & vbCrLf & _
                    "Nein     = Alle Folien" & vbCrLf & _
                    "Abbruch  = Abbrechen", _
                    vbYesNoCancel + vbQuestion, dlgTitle)

    Select Case answer
        Case vbYes:    AskScope = 2   ' Aktuelle Folie
        Case vbNo:     AskScope = 0   ' Alle Folien
        Case vbCancel: AskScope = -1  ' Abbruch
    End Select
End Function


' -----------------------------------------------------------------------
' Zweck:      Gibt SlideRange anhand Scope zurück.
' Parameter:  scope - 0=Alle, 1=Selektierte, 2=Aktuelle Folie
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
