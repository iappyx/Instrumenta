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
    
    Dim reserved(3) As String
    reserved(0) = "slidewidth"
    reserved(1) = "slideheight"
    reserved(2) = "slidecenterx"
    reserved(3) = "slidecentery"
    Dim j As Integer
    For j = 0 To 3
        If varName = reserved(j) Then
            IScr_Log "Line " & lineNum & ": WARNING - """ & varName & """ is read-only"
            Exit Sub
        End If
    Next j
    
    
    Dim valueExpr As String
    valueExpr = Trim(Mid(rest, eqPos + 1))


    If UCase(left(valueExpr, 6)) = "INPUT " Then
        Dim IScr_prompt As String
        IScr_prompt = IScr_ComputeText(Trim(Mid(valueExpr, 7)))
        Dim IScr_inputStr As String
        IScr_inputStr = InputBox(IScr_prompt, "Instrumenta Script")
        If IScr_inputStr = "" Then
            IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = INPUT cancelled / empty"
            Exit Sub
        End If

        If IsNumeric(IScr_inputStr) Then
            IScr_SetVar varName, CDbl(IScr_inputStr)
            IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = " & CDbl(IScr_inputStr) & " (from INPUT)"
        Else
            IScr_SetVarString varName, IScr_inputStr
            IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = """ & IScr_inputStr & """ (from INPUT)"
        End If
        Exit Sub
    End If


    Dim IScr_upperVal As String
    IScr_upperVal = UCase(valueExpr)
    If left(IScr_upperVal, 4) = "GET " Then
        Dim IScr_fromPos As Integer
        IScr_fromPos = InStr(IScr_upperVal, " FROM ")
        If IScr_fromPos = 0 Then
            IScr_Log "Line " & lineNum & ": ERROR - SET VAR GET requires FROM ""shapename"""
            Exit Sub
        End If
        Dim IScr_getProp As String
        IScr_getProp = LCase(Trim(Mid(valueExpr, 5, IScr_fromPos - 5)))
        Dim IScr_getNameExpr As String
        IScr_getNameExpr = Trim(Mid(valueExpr, IScr_fromPos + 6))
        Dim IScr_getShpName As String
        IScr_getShpName = IScr_ComputeText(IScr_getNameExpr)

        Dim IScr_oSlide As Slide
        Set IScr_oSlide = ActiveWindow.View.Slide
        Dim IScr_getShp As shape
        Dim IScr_s As shape
        For Each IScr_s In IScr_oSlide.shapes
            If LCase(IScr_s.name) = LCase(IScr_getShpName) Then
                Set IScr_getShp = IScr_s
                Exit For
            End If
        Next IScr_s
        If IScr_getShp Is Nothing Then
            IScr_Log "Line " & lineNum & ": ERROR - SET VAR GET: shape """ & IScr_getShpName & """ not found"
            Exit Sub
        End If

        Dim IScr_gotNum As Double
        Dim IScr_gotStr As String
        Dim IScr_isStr As Boolean
        IScr_isStr = False
        Select Case IScr_getProp
            Case "position.x":  IScr_gotNum = IScr_getShp.left
            Case "position.y":  IScr_gotNum = IScr_getShp.Top
            Case "width":       IScr_gotNum = IScr_getShp.width
            Case "height":      IScr_gotNum = IScr_getShp.height
            Case "opacity":     IScr_gotNum = (1 - IScr_getShp.Fill.Transparency) * 100
            Case "rotation":    IScr_gotNum = IScr_getShp.rotation
            Case "font.size":
                If IScr_getShp.HasTextFrame Then IScr_gotNum = IScr_getShp.TextFrame.textRange.Font.Size
            Case "name":
                IScr_gotStr = IScr_getShp.name
                IScr_isStr = True
            Case "text":
                If IScr_getShp.HasTextFrame Then IScr_gotStr = IScr_getShp.TextFrame.textRange.text
                IScr_isStr = True
            Case Else
                IScr_Log "Line " & lineNum & ": ERROR - SET VAR GET: unknown property """ & IScr_getProp & """"
                Exit Sub
        End Select

        If IScr_isStr Then
            IScr_SetVarString varName, IScr_gotStr
            IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = GET " & IScr_getProp & " FROM """ & IScr_getShpName & """ -> """ & IScr_gotStr & """"
        Else
            IScr_SetVar varName, IScr_gotNum
            IScr_Log "Line " & lineNum & ": SET VAR " & varName & " = GET " & IScr_getProp & " FROM """ & IScr_getShpName & """ -> " & IScr_gotNum
        End If
        Exit Sub
    End If

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
        Case "font.size", "width", "height", "position.x", "position.y", "opacity", "border.width", "border.radius"
            numVal = IScr_ComputeNumber(valueExpr)
        Case "font.name", "name", "text"
            strVal = IScr_ComputeText(valueExpr)
        Case "font.color", "fill.color", "border.color", "shadow.color", "fill.gradient"
            Dim rawVal As String
            rawVal = Trim(IScr_SubstituteStringVars(valueExpr))
            If left(rawVal, 1) = """" And right(rawVal, 1) = """" Then
                rawVal = Mid(rawVal, 2, Len(rawVal) - 2)
            End If
            strVal = rawVal
        Case "font.bold", "font.italic", "font.underline", "fill.transparent", "border.visible", "border.style", _
             "text.align", "text.valign", "z.order", "shadow", "shadow.offset.x", "shadow.offset.y", _
             "fill.gradient.direction", "connector.style"
            strVal = UCase(Trim(IScr_SubstituteStringVars(Trim(valueExpr))))
   
            If left(strVal, 1) = """" And right(strVal, 1) = """" Then
                strVal = Mid(strVal, 2, Len(strVal) - 2)
            End If
    End Select

    Select Case prop
        Case "font.size":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Size = numVal
        Case "font.bold":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Bold = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.italic":     If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Italic = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.underline":  If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Underline = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.color":      If shp.HasTextFrame Then shp.TextFrame.textRange.Font.color.RGB = IScr_HexToRGB(strVal)
        Case "font.name":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.name = strVal
        Case "text":
            If shp.HasTextFrame Then shp.TextFrame.textRange.text = strVal
        Case "text.align":
            If shp.HasTextFrame Then
                Select Case strVal
                    Case "LEFT":    shp.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignLeft
                    Case "CENTER":  shp.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignCenter
                    Case "RIGHT":   shp.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignRight
                    Case "JUSTIFY": shp.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignJustify
                End Select
            End If
        Case "text.valign":
            If shp.HasTextFrame Then
                Select Case strVal
                    Case "TOP":    shp.TextFrame.VerticalAnchor = msoAnchorTop
                    Case "MIDDLE": shp.TextFrame.VerticalAnchor = msoAnchorMiddle
                    Case "BOTTOM": shp.TextFrame.VerticalAnchor = msoAnchorBottom
                End Select
            End If
        Case "fill.color":      shp.Fill.Solid: shp.Fill.ForeColor.RGB = IScr_HexToRGB(strVal)
        Case "fill.transparent": If (strVal = "TRUE" Or strVal = "1") Then shp.Fill.visible = msoFalse Else shp.Fill.visible = msoTrue
        Case "fill.gradient":
            
            Dim IScr_commaPos As Integer
            IScr_commaPos = InStr(strVal, ",")
            If IScr_commaPos > 0 Then
                Dim IScr_col1 As String
                Dim IScr_col2 As String
                IScr_col1 = Trim(left(strVal, IScr_commaPos - 1))
                IScr_col2 = Trim(Mid(strVal, IScr_commaPos + 1))
                shp.Fill.TwoColorGradient msoGradientHorizontal, 1
                shp.Fill.GradientStops(1).color.RGB = IScr_HexToRGB(IScr_col1)
                shp.Fill.GradientStops(2).color.RGB = IScr_HexToRGB(IScr_col2)
            Else
                IScr_Log "  WARNING - fill.gradient requires two colors: ""#RRGGBB,#RRGGBB"""
            End If
        Case "width":           shp.width = CSng(numVal)
        Case "height":          shp.height = CSng(numVal)
        Case "position.x":      shp.left = CSng(numVal)
        Case "position.y":      shp.Top = CSng(numVal)
        Case "opacity":         shp.Fill.Transparency = 1 - (CSng(numVal) / 100)
        Case "name":            shp.name = strVal
        Case "z.order":
            Select Case strVal
                Case "FRONT":    shp.ZOrder msoBringToFront
                Case "BACK":     shp.ZOrder msoSendToBack
                Case "FORWARD":  shp.ZOrder msoBringForward
                Case "BACKWARD": shp.ZOrder msoSendBackward
            End Select
        Case "shadow":
            shp.Shadow.visible = (strVal = "TRUE" Or strVal = "1")
        Case "shadow.color":
            shp.Shadow.visible = msoTrue
            shp.Shadow.ForeColor.RGB = IScr_HexToRGB(strVal)
        Case "shadow.offset.x":
            shp.Shadow.visible = msoTrue
            shp.Shadow.OffsetX = CSng(IScr_ComputeNumber(valueExpr))
        Case "shadow.offset.y":
            shp.Shadow.visible = msoTrue
            shp.Shadow.OffsetY = CSng(IScr_ComputeNumber(valueExpr))
        Case "border.color":
            shp.line.ForeColor.RGB = IScr_HexToRGB(strVal)
            shp.line.visible = msoTrue
        Case "border.width":
            shp.line.Weight = CSng(numVal)
            shp.line.visible = msoTrue
        Case "border.radius":
            
            If shp.Type = msoAutoShape Then
                If shp.AutoShapeType = msoShapeRoundedRectangle Then

                    Dim IScr_adjVal As Long
                    IScr_adjVal = CLng((numVal / 100) * 50000)
                    If IScr_adjVal < 0 Then IScr_adjVal = 0
                    If IScr_adjVal > 50000 Then IScr_adjVal = 50000
                    shp.Adjustments(1) = IScr_adjVal
                Else
                    IScr_Log "  WARNING - border.radius only applies to ROUNDEDRECTANGLE shapes"
                End If
            End If
        Case "border.visible":
            If (strVal = "TRUE" Or strVal = "1") Then shp.line.visible = msoTrue Else shp.line.visible = msoFalse
        Case "border.style":
            Select Case strVal
                Case "SOLID":    shp.line.DashStyle = msoLineSolid
                Case "DASH":     shp.line.DashStyle = msoLineDash
                Case "DOT":      shp.line.DashStyle = msoLineRoundDot
                Case "DASHDOT":  shp.line.DashStyle = msoLineDashDot
            End Select
        Case "fill.gradient.direction":
        
            If shp.Fill.Type = msoFillGradient Then
                Dim IScr_gs1 As Long
                Dim IScr_gs2 As Long
                IScr_gs1 = shp.Fill.GradientStops(1).color.RGB
                IScr_gs2 = shp.Fill.GradientStops(2).color.RGB
                Select Case strVal
                    Case "HORIZONTAL"
                        shp.Fill.TwoColorGradient msoGradientHorizontal, 1
                    Case "VERTICAL"
                        shp.Fill.TwoColorGradient msoGradientVertical, 1
                    Case "DIAGONAL", "DIAGONAL_UP"
                        shp.Fill.TwoColorGradient msoGradientDiagonalUp, 1
                    Case "DIAGONAL_DOWN"
                        shp.Fill.TwoColorGradient msoGradientDiagonalDown, 1
                End Select
                shp.Fill.GradientStops(1).color.RGB = IScr_gs1
                shp.Fill.GradientStops(2).color.RGB = IScr_gs2
            Else
                IScr_Log "  WARNING - fill.gradient.direction: shape has no gradient fill. Set fill.gradient first."
            End If
        Case "connector.style":
            If shp.ConnectorFormat.Type <> msoConnectorNone Then
                Select Case strVal
                    Case "STRAIGHT": shp.ConnectorFormat.Type = msoConnectorStraight
                    Case "ELBOW":    shp.ConnectorFormat.Type = msoConnectorElbow
                    Case "CURVED":   shp.ConnectorFormat.Type = msoConnectorCurve
                End Select
            Else
                IScr_Log "  WARNING - connector.style: shape """ & shp.name & """ is not a connector"
            End If
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

