Attribute VB_Name = "modBrandCompliance"
Option Explicit

' =============================================================================
' Modul:  modBrandCompliance
' Zweck:  Brand Compliance Checker – prüft alle Shapes einer Präsentation
'         auf Einhaltung von Markenrichtlinien (Schrift, Farben, Größen).
'
' Konfiguration: Infront_BrandConfig.ini im selben Ordner wie die .ppam
' Dateizugriff:  Ausschließlich Open / Line Input / Print # / Close
' Plattform:     Windows und Mac
' =============================================================================

' --- Interne Typen -----------------------------------------------------------

Private Type BrandProfile
    ProfileName     As String
    AllowedFonts()  As String
    FontCount       As Long
    AllowedColors() As Long
    ColorCount      As Long
    ColorTolerance  As Long
    MinFontSizePt   As Double
End Type

' Öffentlich, damit frmBrandCompliance darauf zugreifen kann
Public Type ViolationInfo
    SlideIndex      As Long
    SlideTitle      As String
    ShapePath       As String   ' z.B. "Gruppe 1 > Rechteck 2"
    ViolationType   As String   ' "Font" | "FontSize" | "FillColor" | "LineColor"
    ActualValue     As String
    ExpectedValues  As String
End Type

' Modulweite Zustandsvariablen für die Ergebnisform
Public g_Violations()   As ViolationInfo
Public g_ViolCount      As Long


' =============================================================================
' ÖFFENTLICHE ENTRY-POINTS
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – startet den Brand Compliance Check.
' Parameter:  control - IRibbonControl
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub ShowBrandCheck(control As IRibbonControl)

    Const PROC_NAME As String = "ShowBrandCheck"
    Const DLG_TITLE As String = "Infront Toolkit – Brand Check"

    On Error GoTo ErrHandler

    ' -- Präsentation vorhanden?
    If Presentations.Count = 0 Then
        MsgBox "Keine Präsentation geöffnet.", vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' -- Konfigurationspfad ermitteln
    Dim cfgPath As String
    cfgPath = GetConfigPath()

    If cfgPath = "" Then
        MsgBox "Add-in-Pfad konnte nicht ermittelt werden." & vbCrLf & _
               "Bitte sicherstellen, dass die .ppam korrekt installiert ist.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' -- Konfiguration prüfen / erstellen
    If Dir(cfgPath) = "" Then
        CreateDefaultConfig cfgPath
        MsgBox "Keine Konfigurationsdatei gefunden." & vbCrLf & _
               "Eine Vorlage wurde erstellt unter:" & vbCrLf & cfgPath & vbCrLf & vbCrLf & _
               "Bitte Datei bearbeiten und Brand Check erneut starten.", _
               vbInformation, DLG_TITLE
        Exit Sub
    End If

    ' -- Aktives Profil laden
    Dim activeProfileName As String
    activeProfileName = ReadIniValue(cfgPath, "General", "ActiveProfile", "Default")

    Dim profile As BrandProfile
    If Not LoadProfile(cfgPath, activeProfileName, profile) Then
        MsgBox "Profil '" & activeProfileName & "' konnte nicht geladen werden." & vbCrLf & _
               "Bitte Konfigurationsdatei prüfen: " & cfgPath, _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ' -- Prüfung ausführen
    g_ViolCount = 0
    ReDim g_Violations(1 To 2000)

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        Dim slideTitle As String
        slideTitle = "Folie " & sld.SlideIndex
        On Error Resume Next
        If sld.Shapes.HasTitle Then
            If sld.Shapes.Title.HasTextFrame Then
                slideTitle = slideTitle & ": " & _
                    Left(sld.Shapes.Title.TextFrame.TextRange.Text, 40)
            End If
        End If
        On Error GoTo ErrHandler

        For Each shp In sld.Shapes
            CheckShape shp, profile, sld.SlideIndex, slideTitle, shp.Name
        Next shp
    Next sld

    ' -- Ergebnis anzeigen
    If g_ViolCount = 0 Then
        MsgBox "Keine Brand-Verstöße gefunden." & vbCrLf & _
               "Profil: " & profile.ProfileName, _
               vbInformation, DLG_TITLE
        Exit Sub
    End If

    frmBrandCompliance.InitForm
    frmBrandCompliance.Show vbModeless

    Exit Sub

ErrHandler:
    MsgBox "Fehler in " & PROC_NAME & ": " & Err.Description & _
           " (Nr. " & Err.Number & ")", vbCritical, DLG_TITLE
End Sub


' =============================================================================
' KONFIGURATION
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Gibt vollständigen Pfad zur Konfigurationsdatei zurück.
' Parameter:  (keiner)
' Rückgabe:   Vollständiger Pfad oder "" bei Fehler
' -----------------------------------------------------------------------
Public Function GetConfigPath() As String

    Dim addinPath As String

    ' Primär: ThisPresentation.Path (zuverlässigst in einer .ppam)
    On Error Resume Next
    addinPath = ThisPresentation.Path
    On Error GoTo 0

    ' Fallback: AddIns-Kollektion nach "Infront" durchsuchen
    If addinPath = "" Then
        On Error Resume Next
        Dim addin As AddIn
        For Each addin In Application.AddIns
            If InStr(LCase(addin.FullName), "infront") > 0 Then
                Dim fullName As String
                fullName = addin.FullName
                Dim sep As String
                sep = Application.PathSeparator
                addinPath = Left(fullName, InStrRev(fullName, sep) - 1)
                Exit For
            End If
        Next addin
        On Error GoTo 0
    End If

    If addinPath = "" Then
        GetConfigPath = ""
        Exit Function
    End If

    GetConfigPath = addinPath & Application.PathSeparator & "Infront_BrandConfig.ini"
End Function


' -----------------------------------------------------------------------
' Zweck:      Lädt ein Profil aus der INI-Datei.
' Parameter:  cfgPath     - Pfad zur INI-Datei
'             profileName - Name des Profils (entspricht [Profile.Name])
'             profile     - Ausgabe-Profil-Struct
' Rückgabe:   True bei Erfolg
' -----------------------------------------------------------------------
Private Function LoadProfile(cfgPath As String, profileName As String, _
                             ByRef profile As BrandProfile) As Boolean

    Const PROC_NAME As String = "LoadProfile"

    On Error GoTo LoadError

    Dim section As String
    section = "Profile." & profileName

    profile.ProfileName = ReadIniValue(cfgPath, section, "Name", profileName)
    profile.ColorTolerance = CLng(ReadIniValue(cfgPath, section, "ColorTolerance", "10"))
    profile.MinFontSizePt = CDbl(ReadIniValue(cfgPath, section, "MinFontSizePt", "8"))

    ' Farben einlesen
    Dim colorsRaw As String
    colorsRaw = ReadIniValue(cfgPath, section, "AllowedColors", "")
    ParseColorList colorsRaw, profile.AllowedColors, profile.ColorCount

    ' Schriftarten einlesen
    Dim fontsRaw As String
    fontsRaw = ReadIniValue(cfgPath, section, "AllowedFonts", "")
    ParseFontList fontsRaw, profile.AllowedFonts, profile.FontCount

    LoadProfile = True
    Exit Function

LoadError:
    LoadProfile = False
End Function


' -----------------------------------------------------------------------
' Zweck:      Erstellt eine kommentierte Standard-INI wenn keine vorhanden.
' Parameter:  cfgPath - Zielpfad für neue Datei
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub CreateDefaultConfig(cfgPath As String)

    Dim fileNum As Integer
    On Error GoTo CreateError

    fileNum = FreeFile
    Open cfgPath For Output As #fileNum

    Print #fileNum, "; Infront Toolkit - Brand Compliance Configuration"
    Print #fileNum, "; Erstellt: " & Format(Now, "DD.MM.YYYY HH:MM")
    Print #fileNum, ";"
    Print #fileNum, "; EINRICHTUNG:"
    Print #fileNum, "; 1. AllowedFonts: Kommagetrennte Liste erlaubter Schriftarten"
    Print #fileNum, "; 2. AllowedColors: Kommagetrennte Hex-Werte (#RRGGBB)"
    Print #fileNum, "; 3. ColorTolerance: 0-30 (Abweichung je Farbkanal)"
    Print #fileNum, "; 4. MinFontSizePt: Mindest-Schriftgröße (0 = keine Prüfung)"
    Print #fileNum, ""
    Print #fileNum, "[General]"
    Print #fileNum, "ActiveProfile=Default"
    Print #fileNum, ""
    Print #fileNum, "[Profile.Default]"
    Print #fileNum, "Name=Default"
    Print #fileNum, "AllowedFonts=Calibri,Calibri Light,Arial"
    Print #fileNum, "AllowedColors=#003366,#FFFFFF,#000000"
    Print #fileNum, "ColorTolerance=10"
    Print #fileNum, "MinFontSizePt=8"

    Close #fileNum
    Exit Sub

CreateError:
    On Error Resume Next
    Close #fileNum
End Sub


' =============================================================================
' INI-PARSER (Open/Line Input/Print#/Close – kein FSO)
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Liest einen Wert aus einer INI-Datei.
' Parameter:  filePath   - Pfad zur INI-Datei
'             section    - Abschnittsname (ohne [])
'             key        - Schlüsselname
'             defaultVal - Rückgabewert wenn nicht gefunden
' Rückgabe:   Wert als String
' -----------------------------------------------------------------------
Public Function ReadIniValue(filePath As String, section As String, _
                             key As String, defaultVal As String) As String

    Dim fileNum As Integer
    Dim lineText As String
    Dim inSection As Boolean
    Dim eqPos As Long

    On Error GoTo ReturnDefault

    If Dir(filePath) = "" Then GoTo ReturnDefault

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    inSection = False

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)

        ' Kommentare und Leerzeilen
        If Len(lineText) = 0 Then GoTo NextLine
        If Left(lineText, 1) = ";" Then GoTo NextLine
        If Left(lineText, 1) = "#" Then GoTo NextLine

        ' Abschnitts-Header
        If Left(lineText, 1) = "[" And Right(lineText, 1) = "]" Then
            inSection = (LCase(Mid(lineText, 2, Len(lineText) - 2)) = LCase(section))
            GoTo NextLine
        End If

        ' Key=Value im gesuchten Abschnitt
        If inSection Then
            eqPos = InStr(lineText, "=")
            If eqPos > 1 Then
                If LCase(Trim(Left(lineText, eqPos - 1))) = LCase(key) Then
                    ReadIniValue = Trim(Mid(lineText, eqPos + 1))
                    Close #fileNum
                    Exit Function
                End If
            End If
        End If

NextLine:
    Loop

    Close #fileNum
    ReadIniValue = defaultVal
    Exit Function

ReturnDefault:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    ReadIniValue = defaultVal
End Function


' -----------------------------------------------------------------------
' Zweck:      Schreibt einen Wert in eine INI-Datei (gesamte Datei wird
'             eingelesen, geändert und zurückgeschrieben).
' Parameter:  filePath - Pfad zur INI-Datei
'             section  - Abschnittsname
'             key      - Schlüsselname
'             value    - neuer Wert
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub WriteIniValue(filePath As String, section As String, _
                         key As String, value As String)

    Dim fileNum As Integer
    Dim lines() As String
    Dim lineCount As Long
    Dim lineText As String
    Dim inSection As Boolean
    Dim keyWritten As Boolean
    Dim sectionFound As Boolean
    Dim i As Long

    On Error GoTo WriteError

    ' --- Datei einlesen
    ReDim lines(1 To 2000)
    lineCount = 0

    If Dir(filePath) <> "" Then
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, lineText
            lineCount = lineCount + 1
            lines(lineCount) = lineText
        Loop
        Close #fileNum
    End If

    ' --- In-Memory-Suche und Ersetzen
    Dim newLines() As String
    ReDim newLines(1 To lineCount + 5)
    Dim newCount As Long
    newCount = 0
    inSection = False
    keyWritten = False
    sectionFound = False

    For i = 1 To lineCount
        lineText = Trim(lines(i))

        ' Abschnitts-Header
        If Left(lineText, 1) = "[" And Right(lineText, 1) = "]" Then
            ' Verlassen des gesuchten Abschnitts – Key noch nicht geschrieben?
            If inSection And Not keyWritten Then
                newCount = newCount + 1
                newLines(newCount) = key & "=" & value
                keyWritten = True
            End If
            inSection = (LCase(Mid(lineText, 2, Len(lineText) - 2)) = LCase(section))
            If inSection Then sectionFound = True
        End If

        ' Key im gesuchten Abschnitt ersetzen
        If inSection And Not keyWritten Then
            Dim eqPos As Long
            eqPos = InStr(lineText, "=")
            If eqPos > 1 Then
                If LCase(Trim(Left(lineText, eqPos - 1))) = LCase(key) Then
                    newCount = newCount + 1
                    newLines(newCount) = key & "=" & value
                    keyWritten = True
                    GoTo NextWriteLine
                End If
            End If
        End If

        newCount = newCount + 1
        newLines(newCount) = lines(i)

NextWriteLine:
    Next i

    ' Abschnitt oder Key noch nicht vorhanden?
    If Not sectionFound Then
        newCount = newCount + 1
        newLines(newCount) = ""
        newCount = newCount + 1
        newLines(newCount) = "[" & section & "]"
    End If
    If Not keyWritten Then
        newCount = newCount + 1
        newLines(newCount) = key & "=" & value
    End If

    ' --- Datei zurückschreiben
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    For i = 1 To newCount
        Print #fileNum, newLines(i)
    Next i
    Close #fileNum
    Exit Sub

WriteError:
    On Error Resume Next
    Close #fileNum
End Sub


' =============================================================================
' PARSE-HELFER
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Parst kommagetrennte Hex-Farbliste (#RRGGBB) in Long-Array.
' Parameter:  raw        - roher String aus INI
'             colors()   - Ausgabe-Array
'             colorCount - Anzahl gültiger Einträge
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub ParseColorList(raw As String, ByRef colors() As Long, _
                           ByRef colorCount As Long)

    colorCount = 0
    If Trim(raw) = "" Then
        ReDim colors(0)
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(raw, ",")
    ReDim colors(1 To UBound(parts) + 1)

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim hexVal As String
        hexVal = Trim(parts(i))
        hexVal = Replace(hexVal, "#", "")
        hexVal = Replace(hexVal, " ", "")
        If Len(hexVal) = 6 Then
            On Error Resume Next
            Dim r As Long, g As Long, b As Long
            r = CLng("&H" & Left(hexVal, 2))
            g = CLng("&H" & Mid(hexVal, 3, 2))
            b = CLng("&H" & Right(hexVal, 2))
            If Err.Number = 0 Then
                colorCount = colorCount + 1
                colors(colorCount) = RGB(r, g, b)
            End If
            On Error GoTo 0
        End If
    Next i
End Sub


' -----------------------------------------------------------------------
' Zweck:      Parst kommagetrennte Schriftartenliste in String-Array.
' Parameter:  raw       - roher String aus INI
'             fonts()   - Ausgabe-Array
'             fontCount - Anzahl gültiger Einträge
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub ParseFontList(raw As String, ByRef fonts() As String, _
                          ByRef fontCount As Long)

    fontCount = 0
    If Trim(raw) = "" Then
        ReDim fonts(0)
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(raw, ",")
    ReDim fonts(1 To UBound(parts) + 1)

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim fontName As String
        fontName = Trim(parts(i))
        If Len(fontName) > 0 Then
            fontCount = fontCount + 1
            fonts(fontCount) = fontName
        End If
    Next i
End Sub


' =============================================================================
' SHAPE-PRÜFUNG
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Prüft ein Shape rekursiv (Gruppen werden traversiert,
'             Tabellen vollständig geprüft).
' Parameter:  shp        - zu prüfendes Shape
'             profile    - aktives Brand-Profil
'             slideIdx   - Folien-Index (1-basiert)
'             slideTitle - Anzeigename der Folie
'             shapePath  - Pfad für Anzeige (z.B. "Gruppe > Shape")
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub CheckShape(shp As Shape, ByRef profile As BrandProfile, _
                       slideIdx As Long, slideTitle As String, _
                       shapePath As String)

    On Error Resume Next   ' Defensive gegen gesperrte/ungültige Shapes

    ' Typ ermitteln
    Dim shpType As Long
    shpType = shp.Type

    ' Gruppen rekursiv durchlaufen
    If shpType = msoGroup Then
        Dim gi As Shape
        For Each gi In shp.GroupItems
            CheckShape gi, profile, slideIdx, slideTitle, _
                       shapePath & " > " & gi.Name
        Next gi
        On Error GoTo 0
        Exit Sub
    End If

    ' Tabellen vollständig prüfen
    If shp.HasTable Then
        CheckTable shp, profile, slideIdx, slideTitle, shapePath
        On Error GoTo 0
        Exit Sub
    End If

    ' --- Füllfarbe prüfen (nur bei Solid-Fill)
    If profile.ColorCount > 0 Then
        If shp.Fill.Visible = msoTrue Then
            If shp.Fill.Type = msoFillSolid Then
                Dim fillColor As Long
                fillColor = shp.Fill.ForeColor.RGB
                If Not IsColorAllowed(fillColor, profile) Then
                    AddViolation slideIdx, slideTitle, shapePath, _
                                 "FillColor", _
                                 ColorToHexStr(fillColor), _
                                 AllowedColorsAsString(profile)
                End If
            End If
        End If
    End If

    ' --- Linienfarbe prüfen
    If profile.ColorCount > 0 Then
        If shp.Line.Visible = msoTrue Then
            Dim lineColor As Long
            lineColor = shp.Line.ForeColor.RGB
            If Not IsColorAllowed(lineColor, profile) Then
                AddViolation slideIdx, slideTitle, shapePath, _
                             "LineColor", _
                             ColorToHexStr(lineColor), _
                             AllowedColorsAsString(profile)
            End If
        End If
    End If

    ' --- Text/Font prüfen
    If shp.HasTextFrame Then
        CheckTextFrame shp.TextFrame, profile, slideIdx, slideTitle, shapePath
    End If

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Prüft alle Zellen einer Tabelle auf Schrift und Füllfarbe.
' Parameter:  shp        - Shape mit .HasTable = True
'             profile    - aktives Brand-Profil
'             slideIdx, slideTitle, shapePath - Kontext für Meldungen
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub CheckTable(shp As Shape, ByRef profile As BrandProfile, _
                       slideIdx As Long, slideTitle As String, _
                       shapePath As String)

    On Error Resume Next

    Dim tbl As Table
    Set tbl = shp.Table
    If tbl Is Nothing Then Exit Sub

    Dim r As Long, c As Long
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim cell As cell
            Set cell = tbl.cell(r, c)

            Dim cellPath As String
            cellPath = shapePath & " [Z" & r & "/S" & c & "]"

            ' Zell-Füllfarbe
            If profile.ColorCount > 0 Then
                If cell.Shape.Fill.Visible = msoTrue Then
                    If cell.Shape.Fill.Type = msoFillSolid Then
                        Dim cellFill As Long
                        cellFill = cell.Shape.Fill.ForeColor.RGB
                        If Not IsColorAllowed(cellFill, profile) Then
                            AddViolation slideIdx, slideTitle, cellPath, _
                                         "FillColor", _
                                         ColorToHexStr(cellFill), _
                                         AllowedColorsAsString(profile)
                        End If
                    End If
                End If
            End If

            ' Zell-Text
            If cell.Shape.HasTextFrame Then
                CheckTextFrame cell.Shape.TextFrame, profile, _
                               slideIdx, slideTitle, cellPath
            End If
        Next c
    Next r

    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------
' Zweck:      Prüft Schriftarten und -größen in einem TextFrame.
' Parameter:  tf         - zu prüfendes TextFrame
'             profile    - aktives Brand-Profil
'             slideIdx, slideTitle, shapePath - Kontext
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub CheckTextFrame(tf As TextFrame, ByRef profile As BrandProfile, _
                           slideIdx As Long, slideTitle As String, _
                           shapePath As String)

    On Error Resume Next

    If Not tf.HasText Then Exit Sub

    Dim tr As TextRange
    Set tr = tf.TextRange

    Dim p As Long, rn As Long
    For p = 1 To tr.Paragraphs.Count
        For rn = 1 To tr.Paragraphs(p).Runs.Count

            Dim run As TextRange
            Set run = tr.Paragraphs(p).Runs(rn)

            ' Schriftart prüfen (nur explizit gesetzte)
            Dim fontName As String
            fontName = run.Font.Name
            If fontName <> "" And profile.FontCount > 0 Then
                If Not IsFontAllowed(fontName, profile) Then
                    AddViolation slideIdx, slideTitle, shapePath, _
                                 "Font", fontName, AllowedFontsAsString(profile)
                End If
            End If

            ' Schriftgröße prüfen (nur explizit gesetzte)
            If profile.MinFontSizePt > 0 Then
                Dim fontSize As Double
                fontSize = run.Font.Size
                If fontSize > 0 And fontSize < profile.MinFontSizePt Then
                    AddViolation slideIdx, slideTitle, shapePath, _
                                 "FontSize", _
                                 CStr(fontSize) & " pt", _
                                 "mind. " & CStr(profile.MinFontSizePt) & " pt"
                End If
            End If

        Next rn
    Next p

    On Error GoTo 0
End Sub


' =============================================================================
' PRÜF-HELFER
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Prüft ob eine Farbe innerhalb der Toleranz eines erlaubten Werts liegt.
' Parameter:  testColor - zu prüfende Farbe (RGB Long)
'             profile   - aktives Profil mit AllowedColors und Toleranz
' Rückgabe:   True wenn erlaubt
' -----------------------------------------------------------------------
Private Function IsColorAllowed(testColor As Long, _
                                ByRef profile As BrandProfile) As Boolean

    If profile.ColorCount = 0 Then
        IsColorAllowed = True
        Exit Function
    End If

    Dim i As Long
    For i = 1 To profile.ColorCount
        If ColorMaxChannelDiff(testColor, profile.AllowedColors(i)) _
           <= profile.ColorTolerance Then
            IsColorAllowed = True
            Exit Function
        End If
    Next i

    IsColorAllowed = False
End Function


' -----------------------------------------------------------------------
' Zweck:      Berechnet maximale Abweichung je Farbkanal zwischen zwei Farben.
' Parameter:  c1, c2 - RGB-Farbwerte
' Rückgabe:   Maximale Kanal-Differenz (0–255)
' -----------------------------------------------------------------------
Private Function ColorMaxChannelDiff(c1 As Long, c2 As Long) As Long

    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim dr As Long, dg As Long, db As Long

    r1 = c1 And &HFF:     g1 = (c1 \ &H100) And &HFF: b1 = (c1 \ &H10000) And &HFF
    r2 = c2 And &HFF:     g2 = (c2 \ &H100) And &HFF: b2 = (c2 \ &H10000) And &HFF

    dr = Abs(r1 - r2)
    dg = Abs(g1 - g2)
    db = Abs(b1 - b2)

    Dim maxDiff As Long
    maxDiff = dr
    If dg > maxDiff Then maxDiff = dg
    If db > maxDiff Then maxDiff = db

    ColorMaxChannelDiff = maxDiff
End Function


' -----------------------------------------------------------------------
' Zweck:      Findet die nächste erlaubte Farbe zum gegebenen Wert.
' Parameter:  testColor - Ausgangfarbe
'             profile   - aktives Profil
' Rückgabe:   Nächste erlaubte Farbe (Long)
' -----------------------------------------------------------------------
Public Function NearestAllowedColor(testColor As Long, _
                                    ByRef profile As BrandProfile) As Long

    If profile.ColorCount = 0 Then
        NearestAllowedColor = testColor
        Exit Function
    End If

    Dim bestColor As Long
    Dim bestDiff As Long
    bestColor = profile.AllowedColors(1)
    bestDiff = ColorMaxChannelDiff(testColor, profile.AllowedColors(1))

    Dim i As Long
    For i = 2 To profile.ColorCount
        Dim d As Long
        d = ColorMaxChannelDiff(testColor, profile.AllowedColors(i))
        If d < bestDiff Then
            bestDiff = d
            bestColor = profile.AllowedColors(i)
        End If
    Next i

    NearestAllowedColor = bestColor
End Function


' -----------------------------------------------------------------------
' Zweck:      Prüft ob Schriftart in der erlaubten Liste ist.
' Parameter:  fontName - zu prüfender Schriftname
'             profile  - aktives Profil
' Rückgabe:   True wenn erlaubt
' -----------------------------------------------------------------------
Private Function IsFontAllowed(fontName As String, _
                               ByRef profile As BrandProfile) As Boolean

    If profile.FontCount = 0 Then
        IsFontAllowed = True
        Exit Function
    End If

    Dim i As Long
    For i = 1 To profile.FontCount
        If LCase(Trim(fontName)) = LCase(Trim(profile.AllowedFonts(i))) Then
            IsFontAllowed = True
            Exit Function
        End If
    Next i

    IsFontAllowed = False
End Function


' =============================================================================
' VIOLATIONS-VERWALTUNG
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Fügt einen Verstoß dem globalen Array hinzu.
' Parameter:  slideIdx, slideTitle, shapePath, violType, actual, expected
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Private Sub AddViolation(slideIdx As Long, slideTitle As String, _
                         shapePath As String, violType As String, _
                         actual As String, expected As String)

    g_ViolCount = g_ViolCount + 1

    ' Array bei Bedarf vergrößern
    If g_ViolCount > UBound(g_Violations) Then
        ReDim Preserve g_Violations(1 To g_ViolCount + 500)
    End If

    With g_Violations(g_ViolCount)
        .SlideIndex = slideIdx
        .SlideTitle = slideTitle
        .ShapePath = shapePath
        .ViolationType = violType
        .ActualValue = actual
        .ExpectedValues = expected
    End With
End Sub


' =============================================================================
' CSV-EXPORT
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Exportiert Verstöße als CSV-Datei.
' Parameter:  violCount - Anzahl Einträge in g_Violations
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub ExportViolationsToCSV()

    Const PROC_NAME As String = "ExportViolationsToCSV"
    Const DLG_TITLE As String = "Infront Toolkit – Brand Check"

    On Error GoTo ErrHandler

    If g_ViolCount = 0 Then
        MsgBox "Keine Verstöße zum Exportieren.", vbInformation, DLG_TITLE
        Exit Sub
    End If

    ' Speicherort bestimmen
    Dim savePath As String
    Dim defaultName As String
    defaultName = "Infront_BrandReport_" & Format(Now, "YYYYMMDD_HHMM") & ".csv"

#If Mac Then
    savePath = GetMacSavePath(defaultName)
#Else
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    dlg.InitialFileName = defaultName
    If dlg.Show = -1 Then
        savePath = dlg.SelectedItems(1)
    End If
    Set dlg = Nothing
#End If

    If savePath = "" Then Exit Sub

    ' CSV schreiben
    Dim fileNum As Integer
    fileNum = FreeFile
    Open savePath For Output As #fileNum

    ' Header
    Print #fileNum, "Folie;Folientitel;Shape-Pfad;Typ;Istwert;Sollwert"

    Dim i As Long
    For i = 1 To g_ViolCount
        With g_Violations(i)
            Print #fileNum, _
                .SlideIndex & ";" & _
                CsvEscape(.SlideTitle) & ";" & _
                CsvEscape(.ShapePath) & ";" & _
                CsvEscape(.ViolationType) & ";" & _
                CsvEscape(.ActualValue) & ";" & _
                CsvEscape(.ExpectedValues)
        End With
    Next i

    Close #fileNum

    MsgBox "Export erfolgreich: " & savePath, vbInformation, DLG_TITLE
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "Fehler beim Export: " & Err.Description, vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Behebt einen einzelnen Verstoß automatisch.
'             Font   → ersten erlaubten Font setzen
'             FontSize → auf MinFontSizePt setzen (Best-Effort)
'             FillColor / LineColor → nächste erlaubte Farbe setzen
' Parameter:  slideIndex - Folien-Index (1-basiert)
'             shapePath  - Shape-Name für einfache Shapes (kein Gruppen-Lookup)
'             violType   - "Font" | "FontSize" | "FillColor" | "LineColor"
' Rückgabe:   (keiner)
' Hinweis:    Nur für einfache Top-Level-Shapes; bei Gruppen/Tabellen
'             wird die gesamte Folie neu geprüft werden müssen.
' -----------------------------------------------------------------------
Public Sub FixViolation(slideIndex As Long, shapePath As String, violType As String)

    Const DLG_TITLE As String = "Infront Toolkit – Brand Check"

    On Error GoTo ErrHandler

    If slideIndex < 1 Or slideIndex > ActivePresentation.Slides.Count Then Exit Sub

    Dim sld As Slide
    Set sld = ActivePresentation.Slides(slideIndex)

    ' Konfiguration erneut laden für Zielwerte
    Dim cfgPath As String
    cfgPath = GetConfigPath()
    If cfgPath = "" Then Exit Sub
    If Dir(cfgPath) = "" Then Exit Sub

    Dim activeProfileName As String
    activeProfileName = ReadIniValue(cfgPath, "General", "ActiveProfile", "Default")

    Dim profile As BrandProfile
    If Not LoadProfile(cfgPath, activeProfileName, profile) Then Exit Sub

    ' Shape auf der Folie finden (nur Top-Level, einfacher Name-Match)
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Name = shapePath Or InStr(shapePath, shp.Name) > 0 Then
            Select Case violType
                Case "FillColor"
                    If profile.ColorCount > 0 Then
                        Dim nearFill As Long
                        nearFill = NearestAllowedColor(shp.Fill.ForeColor.RGB, profile)
                        shp.Fill.ForeColor.RGB = nearFill
                    End If
                Case "LineColor"
                    If profile.ColorCount > 0 Then
                        Dim nearLine As Long
                        nearLine = NearestAllowedColor(shp.Line.ForeColor.RGB, profile)
                        shp.Line.ForeColor.RGB = nearLine
                    End If
                Case "Font"
                    If profile.FontCount > 0 And shp.HasTextFrame Then
                        shp.TextFrame.TextRange.Font.Name = profile.AllowedFonts(1)
                    End If
                Case "FontSize"
                    If profile.MinFontSizePt > 0 And shp.HasTextFrame Then
                        Dim p As Long, rn As Long
                        For p = 1 To shp.TextFrame.TextRange.Paragraphs.Count
                            For rn = 1 To shp.TextFrame.TextRange.Paragraphs(p).Runs.Count
                                Dim r As TextRange
                                Set r = shp.TextFrame.TextRange.Paragraphs(p).Runs(rn)
                                If r.Font.Size > 0 And r.Font.Size < profile.MinFontSizePt Then
                                    r.Font.Size = profile.MinFontSizePt
                                End If
                            Next rn
                        Next p
                    End If
            End Select
            Exit For
        End If
    Next shp

    Exit Sub
ErrHandler:
    MsgBox "Fehler beim Beheben (" & violType & "): " & Err.Description, _
           vbExclamation, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Mac-spezifisch: Fragt Speicherpfad per InputBox ab.
'             (kein nativer SaveAs-Dialog über VBA auf Mac ohne AppleScriptTask)
' Parameter:  defaultName - vorgeschlagener Dateiname
' Rückgabe:   Vollständiger Pfad oder "" bei Abbruch
' -----------------------------------------------------------------------
#If Mac Then
Private Function GetMacSavePath(defaultName As String) As String
    Dim homePath As String
    homePath = Environ("HOME")
    Dim suggestedPath As String
    suggestedPath = homePath & "/Desktop/" & defaultName

    Dim userPath As String
    userPath = InputBox("Speicherpfad für CSV-Export:", _
                        "Infront Toolkit – Brand Check", _
                        suggestedPath)
    GetMacSavePath = Trim(userPath)
End Function
#End If


' =============================================================================
' ANZEIGEHILFEN
' =============================================================================

' -----------------------------------------------------------------------
' Zweck:      Formatiert RGB-Long als #RRGGBB-String.
' Parameter:  colorVal - Farbwert
' Rückgabe:   String "#RRGGBB"
' -----------------------------------------------------------------------
Public Function ColorToHexStr(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And &HFF
    g = (colorVal \ &H100) And &HFF
    b = (colorVal \ &H10000) And &HFF
    ColorToHexStr = "#" & Right("00" & Hex(r), 2) & _
                          Right("00" & Hex(g), 2) & _
                          Right("00" & Hex(b), 2)
End Function


' -----------------------------------------------------------------------
' Zweck:      Gibt erlaubte Farben des Profils als lesbaren String aus.
' Parameter:  profile - aktives Profil
' Rückgabe:   Kommagetrennte Hex-Strings
' -----------------------------------------------------------------------
Private Function AllowedColorsAsString(ByRef profile As BrandProfile) As String
    Dim result As String
    Dim i As Long
    For i = 1 To profile.ColorCount
        If i > 1 Then result = result & ", "
        result = result & ColorToHexStr(profile.AllowedColors(i))
    Next i
    AllowedColorsAsString = result
End Function


' -----------------------------------------------------------------------
' Zweck:      Gibt erlaubte Schriftarten des Profils als lesbaren String aus.
' Parameter:  profile - aktives Profil
' Rückgabe:   Kommagetrennte Namen
' -----------------------------------------------------------------------
Private Function AllowedFontsAsString(ByRef profile As BrandProfile) As String
    Dim result As String
    Dim i As Long
    For i = 1 To profile.FontCount
        If i > 1 Then result = result & ", "
        result = result & profile.AllowedFonts(i)
    Next i
    AllowedFontsAsString = result
End Function


' -----------------------------------------------------------------------
' Zweck:      Maskiert Semikolons und Anführungszeichen für CSV-Export.
' Parameter:  s - Eingabestring
' Rückgabe:   CSV-sicherer String
' -----------------------------------------------------------------------
Private Function CsvEscape(s As String) As String
    If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCrLf) > 0 Then
        CsvEscape = """" & Replace(s, """", """""") & """"
    Else
        CsvEscape = s
    End If
End Function
