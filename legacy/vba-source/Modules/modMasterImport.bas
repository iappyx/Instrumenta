Attribute VB_Name = "modMasterImport"
Option Explicit

' =============================================================================
' Modul:  modMasterImport
' Zweck:  Infront Master-Importer – kopiert einen Slide-Master aus einer
'         externen Präsentation in die aktive Präsentation.
'
' Workflow:
'   1. Quelldatei auswählen (.pptx/.pptm/.ppam/.potx)
'   2. "Master laden" → ListBox zeigt verfügbare Masters
'   3. Master auswählen + Optionen setzen
'   4. "Importieren" → Master wird kopiert, optional auf Folien angewendet
'
' Plattform:  Windows und Mac
' Hinweis:    Quelldatei wird ReadOnly + WithWindow:=False geöffnet.
'             Sie wird nach dem Importvorgang ohne Speichern geschlossen.
' =============================================================================

' --- Öffentlicher Options-Typ (von frmMasterImport befüllt) ------------------

Public Type ImportOptions
    ApplyToAllSlides      As Boolean   ' Master auf alle Folien anwenden
    ApplyToSelectedSlides As Boolean   ' Master nur auf selektierte Folien
    RemoveUnusedAfterImport As Boolean ' Ungenutzte Masters danach entfernen
End Type


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – öffnet den Master-Importer.
' Parameter:  control - IRibbonControl
' -----------------------------------------------------------------------
Public Sub ShowMasterImport(control As IRibbonControl)

    Const DLG_TITLE As String = "Infront Toolkit – Master-Importer"

    On Error GoTo ErrHandler

    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    frmMasterImport.InitForm
    frmMasterImport.Show vbModeless

    Exit Sub
ErrHandler:
    MsgBox "Fehler in ShowMasterImport: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Öffnet Quelldatei ReadOnly und liest Master-Namen aus.
' Parameter:  filePath    - Pfad zur Quelldatei
'             masterNames - Ausgabe: Array mit Master-Namen (1-basiert)
'             masterCount - Ausgabe: Anzahl gefundener Masters
' Rückgabe:   True bei Erfolg
' -----------------------------------------------------------------------
Public Function LoadMastersFromFile(filePath As String, _
                                    ByRef masterNames() As String, _
                                    ByRef masterCount As Long) As Boolean
    Const DLG_TITLE As String = "Infront Toolkit – Master-Importer"

    masterCount = 0
    LoadMastersFromFile = False

    If filePath = "" Then Exit Function
    If Dir(filePath) = "" Then
        MsgBox "Datei nicht gefunden: " & vbCrLf & filePath, _
               vbExclamation, DLG_TITLE
        Exit Function
    End If

    Dim srcPres As Presentation
    On Error GoTo OpenError

    Set srcPres = Presentations.Open( _
        FileName:=filePath, _
        ReadOnly:=msoTrue, _
        WithWindow:=msoFalse)

    On Error GoTo CloseAndError

    masterCount = srcPres.SlideMasters.Count
    ReDim masterNames(1 To masterCount)

    Dim i As Long
    For i = 1 To masterCount
        Dim mName As String
        mName = Trim(srcPres.SlideMasters(i).Name)
        If mName = "" Then mName = "Master " & i
        masterNames(i) = mName
    Next i

    srcPres.Close
    Set srcPres = Nothing

    LoadMastersFromFile = True
    Exit Function

OpenError:
    MsgBox "Datei konnte nicht geöffnet werden: " & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "Fehler: " & Err.Description, vbExclamation, DLG_TITLE
    Exit Function

CloseAndError:
    Dim errMsg As String
    errMsg = Err.Description
    On Error Resume Next
    If Not srcPres Is Nothing Then srcPres.Close
    Set srcPres = Nothing
    On Error GoTo 0
    MsgBox "Fehler beim Lesen der Masters: " & errMsg, vbExclamation, DLG_TITLE
End Function


' -----------------------------------------------------------------------
' Zweck:      Importiert einen Master aus der Quelldatei in die aktive
'             Präsentation und wendet ihn optional auf Folien an.
' Parameter:  srcPath     - Pfad zur Quelldatei
'             masterIndex - 1-basierter Index des gewünschten Masters
'             opts        - Import-Optionen
' Rückgabe:   True bei Erfolg
' -----------------------------------------------------------------------
Public Function ImportMaster(srcPath As String, masterIndex As Long, _
                             opts As ImportOptions) As Boolean
    Const DLG_TITLE As String = "Infront Toolkit – Master-Importer"

    ImportMaster = False

    If srcPath = "" Or masterIndex < 1 Then Exit Function

    Dim srcPres As Presentation
    Dim newMaster As Master

    On Error GoTo OpenError

    Set srcPres = Presentations.Open( _
        FileName:=srcPath, _
        ReadOnly:=msoTrue, _
        WithWindow:=msoFalse)

    On Error GoTo CloseAndError

    If masterIndex > srcPres.SlideMasters.Count Then
        MsgBox "Master-Index " & masterIndex & " existiert nicht in der Quelldatei.", _
               vbExclamation, DLG_TITLE
        srcPres.Close
        Exit Function
    End If

    ' Master in aktive Präsentation kopieren
    ' PowerPoint kopiert SlideMaster über das Hinzufügen eines Designs
    Dim srcMaster As Master
    Set srcMaster = srcPres.SlideMasters(masterIndex)

    ' Design-Name des Quell-Masters merken
    Dim designName As String
    designName = Trim(srcMaster.Name)
    If designName = "" Then designName = "Infront Master"

    ' Anzahl Masters vor Import merken
    Dim masterCountBefore As Long
    masterCountBefore = ActivePresentation.SlideMasters.Count

    ' Master über Theme-Datei-Pfad kopieren wenn vorhanden,
    ' sonst über Folienkopie-Trick
    Dim imported As Boolean
    imported = False

    ' Methode: Temporäre Folie aus Quelldatei einfügen, dann Master extrahieren
    If srcPres.Slides.Count = 0 Then
        srcPres.Close
        MsgBox "Die Quelldatei enthält keine Folien – Master kann nicht kopiert werden.", _
               vbExclamation, DLG_TITLE
        Exit Function
    End If

    On Error Resume Next
    srcPres.Slides(1).Copy
    Dim tempSlide As Slide
    ActivePresentation.Slides.Paste
    Set tempSlide = ActivePresentation.Slides(ActivePresentation.Slides.Count)

    If Err.Number = 0 And Not tempSlide Is Nothing Then
        ' Neuer Master wurde durch die eingefügte Folie hinzugefügt
        If ActivePresentation.SlideMasters.Count > masterCountBefore Then
            Set newMaster = ActivePresentation.SlideMasters( _
                ActivePresentation.SlideMasters.Count)
            newMaster.Name = designName
            imported = True
        End If
        ' Temporäre Folie entfernen
        tempSlide.Delete
        Set tempSlide = Nothing
    End If
    Err.Clear
    On Error GoTo CloseAndError

    ' Quelldatei schließen
    srcPres.Close
    Set srcPres = Nothing

    If Not imported Or newMaster Is Nothing Then
        MsgBox "Master konnte nicht importiert werden." & vbCrLf & _
               "Möglicherweise ist der Master bereits vorhanden oder " & _
               "die Quelldatei enthält keine kompatiblen Folien.", _
               vbExclamation, DLG_TITLE
        Exit Function
    End If

    ' --- Optional: Master auf Folien anwenden
    If opts.ApplyToAllSlides Or opts.ApplyToSelectedSlides Then
        ApplyMasterToSlides newMaster, opts
    End If

    ' --- Optional: Ungenutzte Masters entfernen
    If opts.RemoveUnusedAfterImport Then
        CleanUpUnusedMasters
    End If

    ImportMaster = True
    Exit Function

OpenError:
    MsgBox "Datei konnte nicht geöffnet werden: " & vbCrLf & srcPath & vbCrLf & _
           "Fehler: " & Err.Description, vbExclamation, DLG_TITLE
    Exit Function

CloseAndError:
    Dim errMsg As String
    errMsg = Err.Description
    On Error Resume Next
    If Not srcPres Is Nothing Then srcPres.Close
    Set srcPres = Nothing
    On Error GoTo 0
    MsgBox "Fehler beim Importieren: " & errMsg, vbCritical, DLG_TITLE
End Function


' =============================================================================
' HILFS-FUNKTIONEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Wendet den importierten Master auf Folien an.
'             Verwendet erstes Layout des Masters als Standard.
' Parameter:  newMaster - der importierte Master
'             opts      - ApplyToAllSlides / ApplyToSelectedSlides
' -----------------------------------------------------------------------
Private Sub ApplyMasterToSlides(newMaster As Master, opts As ImportOptions)

    On Error Resume Next

    ' Erstes Layout des neuen Masters
    Dim targetLayout As CustomLayout
    If newMaster.CustomLayouts.Count > 0 Then
        Set targetLayout = newMaster.CustomLayouts(1)
    Else
        Exit Sub
    End If

    Dim sld As Slide

    If opts.ApplyToAllSlides Then
        For Each sld In ActivePresentation.Slides
            sld.CustomLayout = targetLayout
        Next sld

    ElseIf opts.ApplyToSelectedSlides Then
        Dim sel As Selection
        Set sel = ActiveWindow.Selection
        If sel.Type = ppSelectionSlides Then
            For Each sld In sel.SlideRange
                sld.CustomLayout = targetLayout
            Next sld
        End If
    End If

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Entfernt alle SlideMasters die von keiner Folie referenziert werden.
' Rückgabe:   Anzahl entfernter Masters
' -----------------------------------------------------------------------
Public Function CleanUpUnusedMasters() As Long

    Dim removed As Long
    removed = 0

    ' Welche Master werden tatsächlich verwendet?
    Dim usedMasters() As Boolean
    Dim masterCount As Long
    masterCount = ActivePresentation.SlideMasters.Count
    ReDim usedMasters(1 To masterCount)

    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        Dim masterIdx As Long
        masterIdx = sld.CustomLayout.Parent.Index  ' Parent = SlideMaster
        If Err.Number = 0 And masterIdx >= 1 And masterIdx <= masterCount Then
            usedMasters(masterIdx) = True
        End If
        Err.Clear
        On Error GoTo 0
    Next sld

    ' Ungenutzte rückwärts löschen
    Dim i As Long
    For i = masterCount To 1 Step -1
        If Not usedMasters(i) Then
            On Error Resume Next
            ActivePresentation.SlideMasters(i).Delete
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    CleanUpUnusedMasters = removed
End Function


' -----------------------------------------------------------------------
' Zweck:      Öffnet plattformgerechten Dateiauswahl-Dialog.
' Rückgabe:   Ausgewählter Dateipfad oder "" bei Abbruch
' -----------------------------------------------------------------------
Public Function BrowseForFile() As String

    Dim result As String

#If Mac Then
    ' Mac: InputBox (kein nativer Datei-Dialog ohne AppleScriptTask-Deployment)
    result = InputBox("Pfad zur Quelldatei eingeben (.pptx/.pptm/.ppam/.potx):", _
                      "Infront Toolkit – Master-Importer", "")
    BrowseForFile = Trim(result)
#Else
    ' Windows: nativer Datei-Dialog
    On Error GoTo FallbackInputBox
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Quelldatei für Master-Import auswählen"
    dlg.Filters.Clear
    dlg.Filters.Add "PowerPoint-Dateien", "*.pptx;*.pptm;*.ppam;*.potx;*.pot"
    dlg.Filters.Add "Alle Dateien", "*.*"
    dlg.AllowMultiSelect = False

    If dlg.Show = -1 Then
        result = dlg.SelectedItems(1)
    End If
    Set dlg = Nothing
    BrowseForFile = result
    Exit Function

FallbackInputBox:
    On Error GoTo 0
    result = InputBox("Pfad zur Quelldatei eingeben (.pptx/.pptm/.ppam/.potx):", _
                      "Infront Toolkit – Master-Importer", "")
    BrowseForFile = Trim(result)
#End If
End Function
