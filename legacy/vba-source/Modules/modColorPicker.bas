Attribute VB_Name = "modColorPicker"
Option Explicit

' =============================================================================
' Modul:  modColorPicker
' Zweck:  Screen Color Picker für Windows und Mac.
'
' Windows: Nutzt Windows-API GetCursorPos / GetDC / GetPixel / ReleaseDC,
'          um die Bildschirmfarbe an der aktuellen Mausposition zu lesen.
'          API-Deklarationen sind identisch zum Muster in ModuleEyedropper.bas
'          (Private-Scope – kein Namenskonflikt).
'          COLORREF und VBA-Long RGB haben dasselbe Byte-Layout
'          (R=low, G=mid, B=high) – keine Konvertierung nötig.
'
' Mac:     Nutzt MacScript("choose color") für den macOS NSColorPanel.
'          HINWEIS: Dies ist ein Farbauswahl-Dialog, kein Screen-Eyedropper.
'          Echtes Screen-Color-Picking ist im PowerPoint-Mac-VBA-Kontext nicht
'          robust ohne installiertes AppleScript-Plugin umsetzbar.
'          AppleScriptTask() wurde bewusst nicht verwendet (erfordert Deployment
'          einer .applescript-Datei nach ~/Library/Application Scripts/
'          com.microsoft.Powerpoint/ – zu aufwändig für ein Add-in).
'          Fallback bei MacScript-Fehler: manuelle Hex-/RGB-Eingabe.
'          Rückgabe: -1 bei Abbruch oder Fehler.
'
' Ergebnis: UserForm frmColorPicker zeigt Farbvorschau, Hex, RGB und
'           Buttons zum Anwenden auf Fill / Line / Font der Selektion.
' =============================================================================

' --- Windows-API-Deklarationen (nur auf Windows aktiv) ----------------------
' Muster identisch zu ModuleEyedropper.bas; Private-Scope verhindert Konflikt.
#If Not Mac Then

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDC Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetPixel Lib "gdi32" _
        (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
#Else
    Private Declare Function GetDC Lib "user32" _
        (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" _
        (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type

#End If
' ---------------------------------------------------------------------------


' -----------------------------------------------------------------------
' Zweck:      Ribbon-Callback – startet den Screen Color Picker, zeigt
'             Ergebnisform und ermöglicht Anwenden auf die Selektion.
' Parameter:  control - IRibbonControl (Ribbon-Callback)
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub ShowColorPicker(control As IRibbonControl)

    Const PROC_NAME As String = "ShowColorPicker"
    Const DLG_TITLE As String = "Infront Toolkit – Color Picker"

    On Error GoTo ErrHandler

    ' -- 1. Farbe aufnehmen
    Dim pickedColor As Long
    pickedColor = PickScreenColor()

    If pickedColor = -1 Then Exit Sub  ' Abbruch durch Nutzer oder Fehler

    ' -- 2. Ergebnisform anzeigen
    frmColorPicker.InitForm pickedColor
    frmColorPicker.Show vbModal

    Exit Sub

ErrHandler:
    MsgBox "Fehler in " & PROC_NAME & ": " & Err.Description & _
           " (Nr. " & Err.Number & ")", vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Plattform-Dispatcher: wählt Windows- oder Mac-Pfad.
' Parameter:  (keiner)
' Rückgabe:   RGB-Farbwert als Long, oder -1 bei Abbruch / Fehler
' -----------------------------------------------------------------------
Private Function PickScreenColor() As Long
#If Mac Then
    PickScreenColor = PickScreenColorMac()
#Else
    PickScreenColor = PickScreenColorWindows()
#End If
End Function


' -----------------------------------------------------------------------
' Zweck:      Windows: Liest Bildschirmfarbe an aktueller Mausposition.
'             Nutzer positioniert Maus und klickt OK in einem Dialog.
' Parameter:  (keiner)
' Rückgabe:   RGB-Farbwert als Long, oder -1 bei Abbruch / Fehler
' -----------------------------------------------------------------------
#If Not Mac Then
Private Function PickScreenColorWindows() As Long

    Const DLG_TITLE As String = "Infront Toolkit – Color Picker"

    ' Anleitung
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "Maus über die gewünschte Farbe positionieren" & vbCrLf & _
        "und dann OK klicken.", _
        vbOKCancel + vbInformation, DLG_TITLE)

    If answer = vbCancel Then
        PickScreenColorWindows = -1
        Exit Function
    End If

    ' Cursor-Position lesen
    Dim pt As POINTAPI
    GetCursorPos pt

    ' Screen-DC holen
#If VBA7 And Win64 Then
    Dim hDC As LongPtr
#Else
    Dim hDC As Long
#End If
    hDC = GetDC(0)

    If hDC = 0 Then
        MsgBox "Fehler: Screen-DC konnte nicht geöffnet werden.", _
               vbExclamation, DLG_TITLE
        PickScreenColorWindows = -1
        Exit Function
    End If

    ' Pixelfarbe lesen und DC sofort freigeben
    Dim colorRef As Long
    colorRef = GetPixel(hDC, pt.x, pt.y)
    ReleaseDC 0, hDC   ' Immer freigeben, direkt nach GetPixel

    ' CLR_INVALID (-1) prüfen
    If colorRef = -1 Then
        PickScreenColorWindows = -1
        Exit Function
    End If

    ' COLORREF und VBA-RGB haben dasselbe Byte-Layout: kein Umbau nötig
    PickScreenColorWindows = colorRef

End Function
#End If


' -----------------------------------------------------------------------
' Zweck:      Mac: Zeigt macOS NSColorPanel via MacScript("choose color").
'             KEIN Screen-Eyedropper – Nutzer wählt Farbe im System-Dialog.
'             Fallback: manuelle Hex-Eingabe wenn MacScript versagt.
' Parameter:  (keiner)
' Rückgabe:   RGB-Farbwert als Long, oder -1 bei Abbruch / Fehler
' -----------------------------------------------------------------------
#If Mac Then
Private Function PickScreenColorMac() As Long

    Const DLG_TITLE As String = "Infront Toolkit – Color Picker"
    Const SCALE As Double = 65535#   ' macOS liefert 0-65535 pro Kanal

    On Error GoTo MacFallback

    ' MacScript("choose color") öffnet macOS NSColorPanel
    Dim asResult As String
    asResult = MacScript("choose color")

    ' Rückgabe-Format: "{r, g, b}" mit Werten 0–65535
    ' Defensiv parsen
    asResult = Trim(asResult)
    If Len(asResult) < 5 Then GoTo MacFallback

    asResult = Replace(asResult, "{", "")
    asResult = Replace(asResult, "}", "")

    Dim parts() As String
    parts = Split(asResult, ",")

    If UBound(parts) < 2 Then GoTo MacFallback

    Dim r As Long, g As Long, b As Long
    If Not IsNumeric(Trim(parts(0))) Then GoTo MacFallback
    If Not IsNumeric(Trim(parts(1))) Then GoTo MacFallback
    If Not IsNumeric(Trim(parts(2))) Then GoTo MacFallback

    ' Skalierung 0-65535 → 0-255 (CLng rundet korrekt)
    r = CLng(CDbl(Trim(parts(0))) / SCALE * 255)
    g = CLng(CDbl(Trim(parts(1))) / SCALE * 255)
    b = CLng(CDbl(Trim(parts(2))) / SCALE * 255)

    ' Auf gültigen Bereich clippen
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255

    PickScreenColorMac = RGB(r, g, b)
    Exit Function

MacFallback:
    ' MacScript nicht verfügbar oder abgebrochen – manuelle Eingabe
    On Error GoTo CancelExit
    Dim hexInput As String
    hexInput = InputBox( _
        "Screen-Farbaufnahme auf Mac nicht verfügbar." & vbCrLf & _
        "Bitte Hex-Farbwert eingeben (z.B. 003366):", _
        DLG_TITLE, "")

    If hexInput = "" Then GoTo CancelExit

    hexInput = Replace(hexInput, "#", "")
    If Len(hexInput) <> 6 Then
        MsgBox "Ungültiger Hex-Wert. Format: RRGGBB (ohne #).", _
               vbExclamation, DLG_TITLE
        GoTo CancelExit
    End If

    Dim rv As Long, gv As Long, bv As Long
    On Error GoTo CancelExit
    rv = CLng("&H" & Left(hexInput, 2))
    gv = CLng("&H" & Mid(hexInput, 3, 2))
    bv = CLng("&H" & Right(hexInput, 2))

    PickScreenColorMac = RGB(rv, gv, bv)
    Exit Function

CancelExit:
    PickScreenColorMac = -1

End Function
#End If


' -----------------------------------------------------------------------
' Zweck:      Wendet eine Farbe auf die ausgewählten Shapes an.
'             Wird vom frmColorPicker aufgerufen.
' Parameter:  pickedColor - RGB-Farbwert (Long)
'             target      - "fill", "line" oder "font"
' Rückgabe:   (keiner)
' -----------------------------------------------------------------------
Public Sub ApplyColorToSelection(pickedColor As Long, target As String)

    Const PROC_NAME As String = "ApplyColorToSelection"
    Const DLG_TITLE As String = "Infront Toolkit – Color Picker"

    On Error GoTo ErrHandler

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Keine Shapes ausgewählt. Bitte zuerst Shapes selektieren.", _
               vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Dim sr As ShapeRange
    Dim shp As Shape
    Dim appliedCount As Long
    Dim skippedCount As Long

    If ActiveWindow.Selection.HasChildShapeRange Then
        Set sr = ActiveWindow.Selection.ChildShapeRange
    Else
        Set sr = ActiveWindow.Selection.ShapeRange
    End If

    appliedCount = 0
    skippedCount = 0

    Dim i As Long
    For i = 1 To sr.Count
        Set shp = sr(i)
        If ApplyColorToShape(shp, pickedColor, target) Then
            appliedCount = appliedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next i

    Dim msg As String
    Select Case target
        Case "fill":  msg = "Füllfarbe"
        Case "line":  msg = "Linienfarbe"
        Case "font":  msg = "Schriftfarbe"
        Case Else:    msg = target
    End Select

    If appliedCount = 0 Then
        MsgBox "Keine Shapes konnten angepasst werden.", _
               vbInformation, DLG_TITLE
    ElseIf skippedCount > 0 Then
        MsgBox appliedCount & " Shape(s) mit " & msg & " angepasst." & vbCrLf & _
               skippedCount & " Shape(s) übersprungen.", _
               vbInformation, DLG_TITLE
    Else
        MsgBox appliedCount & " Shape(s) mit " & msg & " angepasst.", _
               vbInformation, DLG_TITLE
    End If

    Exit Sub

ErrHandler:
    MsgBox "Fehler in " & PROC_NAME & ": " & Err.Description & _
           " (Nr. " & Err.Number & ")", vbCritical, DLG_TITLE
End Sub


' -----------------------------------------------------------------------
' Zweck:      Wendet eine Farbe auf ein einzelnes Shape an.
' Parameter:  shp         - Ziel-Shape
'             pickedColor - RGB-Farbwert
'             target      - "fill", "line" oder "font"
' Rückgabe:   True wenn erfolgreich, False wenn nicht unterstützt
' -----------------------------------------------------------------------
Private Function ApplyColorToShape(shp As Shape, _
                                   pickedColor As Long, _
                                   target As String) As Boolean
    On Error GoTo NotSupported

    Select Case target
        Case "fill"
            shp.Fill.ForeColor.RGB = pickedColor
        Case "line"
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = pickedColor
        Case "font"
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.Font.Color.RGB = pickedColor
            Else
                ApplyColorToShape = False
                Exit Function
            End If
    End Select

    ApplyColorToShape = True
    On Error GoTo 0
    Exit Function

NotSupported:
    ApplyColorToShape = False
    On Error GoTo 0
End Function


' -----------------------------------------------------------------------
' Zweck:      Hilfsfunktion: Formatiert RGB-Wert als Hex-String "#RRGGBB".
' Parameter:  colorVal - RGB-Farbwert als Long
' Rückgabe:   String im Format "#RRGGBB"
' -----------------------------------------------------------------------
Public Function ColorToHex(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And &HFF
    g = (colorVal \ &H100) And &HFF
    b = (colorVal \ &H10000) And &HFF
    ColorToHex = "#" & Right("00" & Hex(r), 2) & _
                       Right("00" & Hex(g), 2) & _
                       Right("00" & Hex(b), 2)
End Function


' -----------------------------------------------------------------------
' Zweck:      Hilfsfunktion: Gibt RGB-Komponenten als String zurück.
' Parameter:  colorVal - RGB-Farbwert als Long
' Rückgabe:   String im Format "R: xxx  G: xxx  B: xxx"
' -----------------------------------------------------------------------
Public Function ColorToRGBString(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And &HFF
    g = (colorVal \ &H100) And &HFF
    b = (colorVal \ &H10000) And &HFF
    ColorToRGBString = "R: " & r & "   G: " & g & "   B: " & b
End Function
