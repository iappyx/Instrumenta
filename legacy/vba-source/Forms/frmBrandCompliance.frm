Attribute VB_Name = "frmBrandCompliance"
Attribute VB_Base = "0{00000000-0000-0000-0000-000000000000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' =============================================================================
' Form:   frmBrandCompliance
' Zweck:  Ergebnisform für den Brand Compliance Checker.
'         Zeigt alle gefundenen Verstöße als Liste, ermöglicht Navigation
'         zum betroffenen Slide, selektives Beheben und CSV-Export.
'
' Benötigte Controls (in VBA-IDE anlegen):
' ---------------------------------------------------------------------------
'  Name           | Typ              | Caption / Verwendung
' ---------------------------------------------------------------------------
'  lstViolations  | ListBox          | Listet Verstöße (ColumnCount = 4)
'  lblSummary     | Label            | z.B. "3 Verstöße auf 2 Folien"
'  btnGoToSlide   | CommandButton    | "Zur Folie"
'  btnFixSelected | CommandButton    | "Auswahl beheben"
'  btnExportCSV   | CommandButton    | "Als CSV exportieren"
'  btnClose       | CommandButton    | "Schließen"
' ---------------------------------------------------------------------------
' Empfohlene Form-Eigenschaften:
'  Caption:     "Infront Toolkit – Brand Compliance Checker"
'  Width:       520 pt,  Height: 340 pt
'  BorderStyle: 1 (Single)
'
' lstViolations – empfohlene Spaltenbreiten (ColumnWidths):
'  "40 pt;60 pt;160 pt;80 pt;80 pt;100 pt"
'  Spalten: Folie | Titel | Shape | Typ | Ist | Erwartet
' =============================================================================

Option Explicit


' -----------------------------------------------------------------------
' Zweck:      Befüllt die ListBox und die Zusammenfassungs-Label.
'             Muss vor Form.Show aufgerufen werden.
' Parameter:  (keiner – liest aus modBrandCompliance.g_Violations)
' -----------------------------------------------------------------------
Public Sub InitForm()

    Dim i           As Long
    Dim slides      As New Collection
    Dim slideKey    As String

    On Error Resume Next

    ' --- Zusammenfassung ---------------------------------------------------
    If modBrandCompliance.g_ViolCount = 0 Then
        lblSummary.Caption = "Keine Verstöße gefunden – Präsentation entspricht dem Markenprofil."
    Else
        ' Eindeutige Folien zählen
        Dim uniqueSlides As Long
        uniqueSlides = 0
        Dim prevSlide As Long
        prevSlide = -1
        For i = 1 To modBrandCompliance.g_ViolCount
            If modBrandCompliance.g_Violations(i).SlideIndex <> prevSlide Then
                uniqueSlides = uniqueSlides + 1
                prevSlide = modBrandCompliance.g_Violations(i).SlideIndex
            End If
        Next i
        lblSummary.Caption = modBrandCompliance.g_ViolCount & " " & _
            IIf(modBrandCompliance.g_ViolCount = 1, "Verstoß", "Verstöße") & _
            " auf " & uniqueSlides & " Folie" & IIf(uniqueSlides = 1, "", "n") & " gefunden."
    End If

    ' --- ListBox befüllen --------------------------------------------------
    lstViolations.Clear

    If modBrandCompliance.g_ViolCount > 0 Then
        lstViolations.ColumnCount = 6

        For i = 1 To modBrandCompliance.g_ViolCount
            With modBrandCompliance.g_Violations(i)
                lstViolations.AddItem CStr(.SlideIndex)
                lstViolations.List(i - 1, 1) = .SlideTitle
                lstViolations.List(i - 1, 2) = .ShapePath
                lstViolations.List(i - 1, 3) = .ViolationType
                lstViolations.List(i - 1, 4) = .ActualValue
                lstViolations.List(i - 1, 5) = .ExpectedValues
            End With
        Next i

        ' Ersten Eintrag selektieren
        lstViolations.ListIndex = 0
    End If

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Navigiert zur Folie des selektierten Verstoßes.
' -----------------------------------------------------------------------
Private Sub btnGoToSlide_Click()

    Dim idx         As Long
    Dim slideIndex  As Long

    On Error GoTo ErrHandler

    idx = lstViolations.ListIndex
    If idx < 0 Then
        MsgBox "Bitte zuerst einen Verstoß auswählen.", vbInformation, "Infront Toolkit"
        Exit Sub
    End If

    slideIndex = CLng(lstViolations.List(idx, 0))

    If slideIndex < 1 Or slideIndex > ActivePresentation.Slides.Count Then
        MsgBox "Folie " & slideIndex & " nicht gefunden.", vbExclamation, "Infront Toolkit"
        Exit Sub
    End If

    ActivePresentation.Slides(slideIndex).Select
    ActiveWindow.View.GotoSlide slideIndex

    Exit Sub
ErrHandler:
    MsgBox "Fehler beim Navigieren zur Folie: " & Err.Description, vbExclamation, "Infront Toolkit"
End Sub


' -----------------------------------------------------------------------
' Zweck:      Behebt den selektierten Verstoß automatisch:
'             Font   → setzt ersten erlaubten Font
'             FontSize → setzt auf MinFontSizePt
'             FillColor / LineColor → setzt auf nächste erlaubte Farbe
' -----------------------------------------------------------------------
Private Sub btnFixSelected_Click()

    Dim idx         As Long
    Dim violType    As String
    Dim slideIndex  As Long
    Dim shapePath   As String

    On Error GoTo ErrHandler

    idx = lstViolations.ListIndex
    If idx < 0 Then
        MsgBox "Bitte zuerst einen Verstoß auswählen.", vbInformation, "Infront Toolkit"
        Exit Sub
    End If

    slideIndex = CLng(lstViolations.List(idx, 0))
    shapePath  = lstViolations.List(idx, 2)
    violType   = lstViolations.List(idx, 3)

    ' Folie navigieren
    If slideIndex >= 1 And slideIndex <= ActivePresentation.Slides.Count Then
        ActiveWindow.View.GotoSlide slideIndex
    End If

    modBrandCompliance.FixViolation slideIndex, shapePath, violType

    ' Verstoß aus Liste entfernen
    lstViolations.RemoveItem idx

    ' Zähler aktualisieren
    modBrandCompliance.g_ViolCount = modBrandCompliance.g_ViolCount - 1

    If modBrandCompliance.g_ViolCount <= 0 Then
        lblSummary.Caption = "Alle Verstöße behoben."
    Else
        lblSummary.Caption = modBrandCompliance.g_ViolCount & " Verstoß" & _
            IIf(modBrandCompliance.g_ViolCount = 1, "", "verstöße") & " verbleibend."
    End If

    Exit Sub
ErrHandler:
    MsgBox "Fehler beim Beheben des Verstoßes: " & Err.Description, vbExclamation, "Infront Toolkit"
End Sub


' -----------------------------------------------------------------------
' Zweck:      Exportiert alle Verstöße als CSV-Datei.
' -----------------------------------------------------------------------
Private Sub btnExportCSV_Click()
    modBrandCompliance.ExportViolationsToCSV
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form ohne weitere Aktionen.
' -----------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub


' -----------------------------------------------------------------------
' Zweck:      Schließt die Form sauber bei X-Button oder Alt+F4.
' -----------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = False
    End If
End Sub
