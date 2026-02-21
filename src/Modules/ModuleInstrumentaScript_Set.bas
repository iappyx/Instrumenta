Attribute VB_Name = "ModuleInstrumentaScript_Set"
'MIT License

'Copyright (c) 2021 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Public Sub IScr_ApplySetVar(line As String, lineNum As Integer)
    
    Dim rest As String
    rest = Trim(Mid(line, 8))

    Dim eqPos As Integer
    eqPos = InStr(rest, "=")
    If eqPos = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - SET VAR requires = sign"
        Exit Sub
    End If

    Dim varName As String
    varName = LCase(Trim(left(rest, eqPos - 1)))
    Dim valueExpr As String
    valueExpr = Trim(Mid(rest, eqPos + 1))

    If InStr(valueExpr, """") > 0 Then
        Dim strVal As String
        strVal = IScr_ComputeText(valueExpr)
        IScr_SetVarString varName, strVal
        IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = """ & strVal & """"
    Else
        Dim numVal As Double
        numVal = IScr_ComputeNumber(valueExpr)
        IScr_SetVar varName, numVal
        IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = " & numVal
    End If
End Sub


Public Function IScr_IsIfCommand(upperLine As String) As Boolean
    IScr_IsIfCommand = (left(upperLine, 3) = "IF ")
End Function




Public Sub IScr_SetVar(varName As String, value As Double)
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varNames(i) = varName Then
            IScr_varValues(i) = value
            IScr_varIsString(i) = False
            Exit Sub
        End If
    Next i
    ReDim Preserve IScr_varNames(IScr_varCount)
    ReDim Preserve IScr_varValues(IScr_varCount)
    ReDim Preserve IScr_varStrValues(IScr_varCount)
    ReDim Preserve IScr_varIsString(IScr_varCount)
    IScr_varNames(IScr_varCount) = varName
    IScr_varValues(IScr_varCount) = value
    IScr_varIsString(IScr_varCount) = False
    IScr_varCount = IScr_varCount + 1
End Sub

Public Sub IScr_SetVarString(varName As String, value As String)
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varNames(i) = varName Then
            IScr_varStrValues(i) = value
            IScr_varIsString(i) = True
            Exit Sub
        End If
    Next i
    ReDim Preserve IScr_varNames(IScr_varCount)
    ReDim Preserve IScr_varValues(IScr_varCount)
    ReDim Preserve IScr_varStrValues(IScr_varCount)
    ReDim Preserve IScr_varIsString(IScr_varCount)
    IScr_varNames(IScr_varCount) = varName
    IScr_varStrValues(IScr_varCount) = value
    IScr_varIsString(IScr_varCount) = True
    IScr_varCount = IScr_varCount + 1
End Sub

Public Function IScr_GetVarString(varName As String) As String
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varNames(i) = LCase(varName) Then
            If IScr_varIsString(i) Then
                IScr_GetVarString = IScr_varStrValues(i)
            Else
                If IScr_varValues(i) = Int(IScr_varValues(i)) Then
                    IScr_GetVarString = CStr(CLng(IScr_varValues(i)))
                Else
                    IScr_GetVarString = CStr(IScr_varValues(i))
                End If
            End If
            Exit Function
        End If
    Next i
    IScr_GetVarString = ""
End Function

Public Function IScr_GetVar(varName As String) As Double
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varNames(i) = LCase(varName) Then
            IScr_GetVar = IScr_varValues(i)
            Exit Function
        End If
    Next i
    IScr_GetVar = 0
End Function

Public Function IScr_VarExists(varName As String) As Boolean
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varNames(i) = LCase(varName) Then
            IScr_VarExists = True
            Exit Function
        End If
    Next i
    IScr_VarExists = False
End Function

Public Sub IScr_ApplySet(shapes As Collection, line As String, lineNum As Integer)
    Dim rest As String
    rest = Trim(Mid(line, 4))

    Dim eqPos As Integer
    eqPos = InStr(rest, "=")
    If eqPos = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - SET requires = sign"
        Exit Sub
    End If

    Dim prop As String
    Dim valueExpr As String
    prop = LCase(Trim(left(rest, eqPos - 1)))
    valueExpr = Trim(Mid(rest, eqPos + 1))

    Dim shp As shape
    Dim successCount As Integer
    successCount = 0

    For Each shp In shapes
        If IScr_ApplyProperty(shp, prop, valueExpr, lineNum) Then successCount = successCount + 1
    Next shp

    IScr_Log "Line " & lineNum & ": SET " & prop & " = " & valueExpr & " -> applied to " & successCount & " shape(s)"
End Sub

Public Function IScr_ApplyProperty(shp As shape, prop As String, valueExpr As String, lineNum As Integer) As Boolean
    IScr_ApplyProperty = True
    On Error GoTo Failed

    Dim numVal As Double
    Dim strVal As String

    Select Case prop
        Case "font.size", "width", "height", "position.x", "position.y", "opacity", "border.width"
            numVal = IScr_ComputeNumber(valueExpr)
        Case "font.name", "name"
            strVal = IScr_ComputeText(valueExpr)
      Case "font.color", "fill.color", "border.color"
            Dim rawVal As String
            rawVal = Trim(IScr_SubstituteStringVars(valueExpr))
            If left(rawVal, 1) = """" And right(rawVal, 1) = """" Then
                rawVal = Mid(rawVal, 2, Len(rawVal) - 2)
            End If
            strVal = rawVal
        Case "font.bold", "font.italic", "font.underline", "fill.transparent", "border.visible", "border.style"
            strVal = UCase(Trim(valueExpr))
    End Select

    Select Case prop
        Case "font.size":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Size = numVal
        Case "font.bold":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Bold = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.italic":     If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Italic = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.underline":  If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Underline = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.color":      If shp.HasTextFrame Then shp.TextFrame.textRange.Font.color.RGB = IScr_HexToRGB(strVal)
        Case "font.name":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.name = strVal
        Case "fill.color":      shp.Fill.Solid: shp.Fill.ForeColor.RGB = IScr_HexToRGB(strVal)
        Case "fill.transparent": If (strVal = "TRUE" Or strVal = "1") Then shp.Fill.visible = msoFalse Else shp.Fill.visible = msoTrue
        Case "width":           shp.width = CSng(numVal)
        Case "height":          shp.height = CSng(numVal)
        Case "position.x":      shp.left = CSng(numVal)
        Case "position.y":      shp.Top = CSng(numVal)
        Case "opacity":         shp.Fill.Transparency = 1 - (CSng(numVal) / 100)
        Case "name":            shp.name = strVal
        Case "border.color":
            shp.line.ForeColor.RGB = IScr_HexToRGB(strVal)
            shp.line.visible = msoTrue
        Case "border.width":
            shp.line.Weight = CSng(numVal)
            shp.line.visible = msoTrue
        Case "border.visible":
            If (strVal = "TRUE" Or strVal = "1") Then shp.line.visible = msoTrue Else shp.line.visible = msoFalse
        Case "border.style":
            Select Case strVal
                Case "SOLID":    shp.line.DashStyle = msoLineSolid
                Case "DASH":     shp.line.DashStyle = msoLineDash
                Case "DOT":      shp.line.DashStyle = msoLineRoundDot
                Case "DASHDOT":  shp.line.DashStyle = msoLineDashDot
            End Select
        Case Else
            IScr_Log "  WARNING - Unknown property: " & prop
            IScr_ApplyProperty = False
    End Select
    Exit Function
Failed:
    IScr_Log "  ERROR - Could not set " & prop & " on """ & shp.name & """: " & Err.Description
    IScr_ApplyProperty = False
End Function

Public Function IScr_HexToRGB(hexColor As String) As Long
    Dim h As String
    h = Trim(Replace(hexColor, "#", ""))
    IScr_HexToRGB = RGB(CLng("&H" & left(h, 2)), CLng("&H" & Mid(h, 3, 2)), CLng("&H" & right(h, 2)))
End Function


