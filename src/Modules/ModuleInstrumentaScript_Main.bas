Attribute VB_Name = "ModuleInstrumentaScript_Main"
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

Option Explicit

Public IScr_ScriptLog As Collection

Public IScr_varNames() As String
Public IScr_varValues() As Double
Public IScr_varStrValues() As String
Public IScr_varIsString() As Boolean
Public IScr_varCount As Integer

Public IScr_breakFlag As Boolean
Public IScr_insertCounter As Integer
Public IScr_exprTokens() As String
Public IScr_exprPos As Integer

Public Sub ShowScriptEditor()
    ScriptEditorForm.Show vbModeless
End Sub


Public Sub RunInstrumentaScript(scriptText As String)
    Set IScr_ScriptLog = New Collection
    IScr_insertCounter = 0
    IScr_breakFlag = False
    IScr_varCount = 0
    ReDim IScr_varNames(0)
    ReDim IScr_varValues(0)
    ReDim IScr_varStrValues(0)
    ReDim IScr_varIsString(0)

    Dim rawLines() As String
    If InStr(scriptText, vbLf) > 0 Then
        rawLines = Split(scriptText, vbLf)
    Else
        rawLines = Split(scriptText, vbCr)
    End If
    
    Dim lines() As String
    ReDim lines(0 To UBound(rawLines))
    Dim i As Integer
    For i = 0 To UBound(rawLines)
        lines(i) = Trim(Replace(rawLines(i), vbCr, ""))
    Next i

    Dim lineNum As Integer
    lineNum = 0
    Dim selectedShapes As Collection
    Set selectedShapes = New Collection

    IScr_RunBlock lines, 0, UBound(lines), lineNum, selectedShapes

    IScr_Log "---"
    IScr_Log "Done."
End Sub

Public Function IScr_RunBlock(lines() As String, startIdx As Integer, endIdx As Integer, _
                               ByRef lineNum As Integer, ByRef selectedShapes As Collection) As Integer
    Dim i As Integer
    i = startIdx

    Do While i <= endIdx
        If IScr_breakFlag Then Exit Do

        Dim line As String
        line = lines(i)
        lineNum = i + 1

        If Len(Trim(line)) = 0 Or left(Trim(line), 1) = "#" Then
            i = i + 1
            GoTo NextIteration
        End If

        Dim upperLine As String
        upperLine = UCase(Trim(line))


        If left(upperLine, 6) = "SELECT" Then
            Set selectedShapes = IScr_ExecuteSelect(line, lineNum)
            IScr_SyncSelectionToPowerPoint selectedShapes, lineNum
            IScr_Log "Line " & lineNum & ": Selected " & selectedShapes.count & " shape(s)"

        
        ElseIf upperLine = "USE SELECTION" Then
            Set selectedShapes = IScr_ExecuteUseSelection(lineNum)
            IScr_Log "Line " & lineNum & ": Using PowerPoint selection - " & selectedShapes.count & " shape(s)"

        
        ElseIf left(upperLine, 6) = "INSERT" Then
            Dim newShape As shape
            Set newShape = IScr_ExecuteInsert(line, lineNum)
            If Not newShape Is Nothing Then
                Set selectedShapes = New Collection
                selectedShapes.Add newShape
                IScr_SyncSelectionToPowerPoint selectedShapes, lineNum
                IScr_Log "Line " & lineNum & ": Inserted """ & newShape.name & """ - now working set"
            Else
                Set selectedShapes = New Collection
            End If

        
        ElseIf left(upperLine, 6) = "DELETE" Then
            IScr_ExecuteDelete line, lineNum
            Set selectedShapes = New Collection

        ElseIf left(upperLine, 7) = "SET VAR" Then
            IScr_ApplySetVar line, lineNum
        
        ElseIf left(upperLine, 3) = "SET" Then
            If selectedShapes.count = 0 Then
                IScr_Log "Line " & lineNum & ": WARNING - SET called but no shapes in working set"
            Else
                IScr_ApplySet selectedShapes, line, lineNum
            End If

        
        ElseIf left(upperLine, 4) = "CALL" Then
            IScr_InvokeCommand line, lineNum

        ElseIf left(upperLine, 5) = "GROUP" Then
            Set selectedShapes = IScr_ExecuteGroup(line, selectedShapes, lineNum)

        ElseIf left(upperLine, 6) = "ROTATE" Then
            IScr_ExecuteRotate line, selectedShapes, lineNum

        ElseIf left(upperLine, 6) = "REPEAT" Then
           
            Dim repeatEnd As Integer
            repeatEnd = IScr_FindBlockEnd(lines, i, "REPEAT", "END REPEAT")
            If repeatEnd = -1 Then
                IScr_Log "Line " & lineNum & ": ERROR - No matching END REPEAT found"
                i = endIdx + 1
                GoTo NextIteration
            End If
            IScr_ExecuteRepeat lines, i, repeatEnd, lineNum, selectedShapes
            i = repeatEnd + 1
            GoTo NextIteration

        ElseIf left(upperLine, 2) = "IF" And left(upperLine, 5) <> "IF i " Or _
               (left(upperLine, 2) = "IF" And IScr_IsIfCommand(upperLine)) Then
            Dim ifEnd As Integer
            ifEnd = IScr_FindBlockEnd(lines, i, "IF", "END IF")
            If ifEnd = -1 Then
                IScr_Log "Line " & lineNum & ": ERROR - No matching END IF found"
                i = endIdx + 1
                GoTo NextIteration
            End If
            IScr_ExecuteIf lines, i, ifEnd, lineNum, selectedShapes
            i = ifEnd + 1
            GoTo NextIteration

        ElseIf upperLine = "BREAK" Then
            IScr_Log "Line " & lineNum & ": BREAK"
            IScr_breakFlag = True
            Exit Do


        ElseIf upperLine = "END REPEAT" Or upperLine = "END IF" Or _
               left(upperLine, 4) = "ELSE" Then

            i = i + 1
            GoTo NextIteration

        Else
            IScr_Log "Line " & lineNum & ": ERROR - Unknown command: " & Trim(line)
        End If

        i = i + 1
NextIteration:
    Loop

    IScr_RunBlock = i - 1
End Function

Public Sub IScr_ExecuteRepeat(lines() As String, repeatLine As Integer, endLine As Integer, _
                           ByRef lineNum As Integer, ByRef selectedShapes As Collection)
    Dim headerLine As String
    headerLine = UCase(Trim(lines(repeatLine)))


    Dim parts As String
    parts = Trim(Mid(headerLine, 7))

    Dim countVal As Double
    Dim asPos As Integer
    asPos = InStr(parts, " AS ")
    If asPos = 0 Then
        IScr_Log "Line " & (repeatLine + 1) & ": ERROR - REPEAT requires AS <variable>"
        Exit Sub
    End If
    countVal = IScr_ComputeNumber(Trim(left(parts, asPos - 1)))

    Dim afterAs As String
    afterAs = Trim(Mid(parts, asPos + 4))
    Dim varName As String
    Dim spPos As Integer
    spPos = InStr(afterAs, " ")
    If spPos > 0 Then
        varName = LCase(Trim(left(afterAs, spPos - 1)))
    Else
        varName = LCase(Trim(afterAs))
    End If

    Dim startVal As Double
    startVal = 0
    Dim fromPos As Integer
    fromPos = InStr(parts, " FROM ")
    If fromPos > 0 Then
        Dim afterFrom As String
        afterFrom = Trim(Mid(parts, fromPos + 6))
        Dim stepPos2 As Integer
        stepPos2 = InStr(afterFrom, " STEP ")
        If stepPos2 > 0 Then
            startVal = IScr_ComputeNumber(Trim(left(afterFrom, stepPos2 - 1)))
        Else
            startVal = IScr_ComputeNumber(Trim(afterFrom))
        End If
    End If

    Dim stepVal As Double
    stepVal = 1
    Dim stepPos As Integer
    stepPos = InStr(parts, " STEP ")
    If stepPos > 0 Then
        stepVal = IScr_ComputeNumber(Trim(Mid(parts, stepPos + 6)))
    End If

    Dim bodyStart As Integer
    Dim bodyEnd As Integer
    bodyStart = repeatLine + 1
    bodyEnd = endLine - 1

    Dim loopVal As Double
    Dim iteration As Long
    For iteration = 0 To countVal - 1
        loopVal = startVal + iteration * stepVal
        IScr_SetVar varName, loopVal
        IScr_breakFlag = False
        IScr_RunBlock lines, bodyStart, bodyEnd, lineNum, selectedShapes
        If IScr_breakFlag Then
            IScr_breakFlag = False
            Exit For
        End If
    Next iteration
End Sub


Public Sub IScr_ExecuteIf(lines() As String, ifLine As Integer, endIfLine As Integer, _
                       ByRef lineNum As Integer, ByRef selectedShapes As Collection)

    Dim branchStarts() As Integer
    Dim branchEnds() As Integer
    Dim branchConds() As String
    Dim branchCount As Integer
    branchCount = 0
    ReDim branchStarts(0)
    ReDim branchEnds(0)
    ReDim branchConds(0)

    Dim i As Integer
    i = ifLine
    Dim depth As Integer
    depth = 0

    Do While i <= endIfLine
        Dim uLine As String
        uLine = UCase(Trim(lines(i)))

        If left(uLine, 6) = "REPEAT" Or (left(uLine, 2) = "IF" And i <> ifLine And depth = 0) Then
            depth = depth + 1
        ElseIf uLine = "END REPEAT" Or uLine = "END IF" Then
            If depth > 0 Then depth = depth - 1
        End If

        If depth = 0 Then
            If i = ifLine Then
                
                branchCount = branchCount + 1
                ReDim Preserve branchStarts(branchCount - 1)
                ReDim Preserve branchEnds(branchCount - 1)
                ReDim Preserve branchConds(branchCount - 1)
                branchStarts(branchCount - 1) = i + 1
                branchConds(branchCount - 1) = Trim(Mid(Trim(lines(i)), 3))
            ElseIf left(uLine, 7) = "ELSE IF" Then
               
                branchEnds(branchCount - 1) = i - 1
                
                branchCount = branchCount + 1
                ReDim Preserve branchStarts(branchCount - 1)
                ReDim Preserve branchEnds(branchCount - 1)
                ReDim Preserve branchConds(branchCount - 1)
                branchStarts(branchCount - 1) = i + 1
                branchConds(branchCount - 1) = Trim(Mid(Trim(lines(i)), 8))
            ElseIf uLine = "ELSE" Then
               
                branchEnds(branchCount - 1) = i - 1
               
                branchCount = branchCount + 1
                ReDim Preserve branchStarts(branchCount - 1)
                ReDim Preserve branchEnds(branchCount - 1)
                ReDim Preserve branchConds(branchCount - 1)
                branchStarts(branchCount - 1) = i + 1
                branchConds(branchCount - 1) = "TRUE"
            ElseIf uLine = "END IF" Then
                branchEnds(branchCount - 1) = i - 1
            End If
        End If
        i = i + 1
    Loop

    Dim b As Integer
    For b = 0 To branchCount - 1
        Dim condResult As Boolean
        If branchConds(b) = "TRUE" Then
            condResult = True
        Else
            condResult = IScr_EvalCondition(branchConds(b))
        End If

        If condResult Then
            IScr_RunBlock lines, branchStarts(b), branchEnds(b), lineNum, selectedShapes
            Exit For
        End If
    Next b
End Sub


Public Function IScr_FindBlockEnd(lines() As String, startIdx As Integer, _
                               blockStart As String, blockEnd As String) As Integer
    Dim depth As Integer
    depth = 0
    Dim i As Integer

    For i = startIdx To UBound(lines)
        Dim uLine As String
        uLine = UCase(Trim(lines(i)))

        If left(uLine, Len(blockStart)) = blockStart And i > startIdx Then
            If Not (blockStart = "IF" And left(uLine, 7) = "ELSE IF") Then
                depth = depth + 1
            End If
        End If

        If uLine = blockEnd Then
            If depth = 0 Then
                IScr_FindBlockEnd = i
                Exit Function
            End If
            depth = depth - 1
        End If
    Next i

    IScr_FindBlockEnd = -1
End Function

Public Function IScr_ExecuteUseSelection(lineNum As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    On Error GoTo Failed
    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    If sel.Type = ppSelectionShapes Then
        Dim i As Integer
        For i = 1 To sel.ShapeRange.count
            result.Add sel.ShapeRange(i)
        Next i
    ElseIf sel.Type = ppSelectionText Then
        result.Add sel.ShapeRange(1)
    Else
        IScr_Log "Line " & lineNum & ": WARNING - No shapes in current PowerPoint selection"
    End If

    Set IScr_ExecuteUseSelection = result
    Exit Function
Failed:
    IScr_Log "Line " & lineNum & ": ERROR - Could not read PowerPoint selection"
    Set IScr_ExecuteUseSelection = result
End Function


Public Function IScr_ExecuteSelect(line As String, lineNum As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim upperLine As String
    upperLine = UCase(Trim(line))

    Dim oSlide As Slide
    On Error Resume Next
    Set oSlide = ActiveWindow.View.Slide
    On Error GoTo 0

    If oSlide Is Nothing Then
        IScr_Log "Line " & lineNum & ": ERROR - No active slide"
        Set IScr_ExecuteSelect = result
        Exit Function
    End If

    If upperLine = "SELECT ALL" Then
        Dim shp As shape
        For Each shp In oSlide.shapes
            result.Add shp
        Next shp
        Set IScr_ExecuteSelect = result
        Exit Function
    End If

    If InStr(upperLine, "WHERE") = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - Expected ALL or WHERE after SELECT"
        Set IScr_ExecuteSelect = result
        Exit Function
    End If

    Dim rawCriteria As String
    rawCriteria = Trim(Mid(line, InStr(upperLine, "WHERE") + 5))
    Dim criteria As String
    criteria = IScr_ResolveSelectCriteria(rawCriteria)

    For Each shp In oSlide.shapes
        If IScr_ShapeMatchesCriteria(shp, criteria) Then
            result.Add shp
        End If
    Next shp

    Set IScr_ExecuteSelect = result
End Function

Public Function IScr_ExecuteInsert(line As String, lineNum As Integer) As shape
    Set IScr_ExecuteInsert = Nothing

    Dim upperLine As String
    upperLine = UCase(Trim(line))

   
    Dim shapeType As String
    Dim shapeKeywords(29) As String
    shapeKeywords(0) = "ROUNDEDRECTANGLE"
    shapeKeywords(1) = "RECTANGLE"
    shapeKeywords(2) = "TEXTBOX"
    shapeKeywords(3) = "OVAL"
    shapeKeywords(4) = "RIGHTTRIANGLE"
    shapeKeywords(5) = "TRIANGLE"
    shapeKeywords(6) = "DIAMOND"
    shapeKeywords(7) = "PARALLELOGRAM"
    shapeKeywords(8) = "TRAPEZOID"
    shapeKeywords(9) = "HEXAGON"
    shapeKeywords(10) = "PENTAGON_ARROW"
    shapeKeywords(11) = "PENTAGON"
    shapeKeywords(12) = "OCTAGON"
    shapeKeywords(13) = "ARROWLEFTRIGHT"
    shapeKeywords(14) = "ARROWRIGHT"
    shapeKeywords(15) = "ARROWLEFT"
    shapeKeywords(16) = "ARROWUP"
    shapeKeywords(17) = "ARROWDOWN"
    shapeKeywords(18) = "CHEVRON"
    shapeKeywords(19) = "CIRCULARRIGHTARROW"
    shapeKeywords(20) = "FLOWCHART_PROCESS"
    shapeKeywords(21) = "FLOWCHART_DECISION"
    shapeKeywords(22) = "FLOWCHART_TERMINATOR"
    shapeKeywords(23) = "FLOWCHART_DATA"
    shapeKeywords(24) = "FLOWCHART_DOCUMENT"
    shapeKeywords(25) = "FLOWCHART_CONNECTOR"
    shapeKeywords(26) = "CALLOUT_RECT"
    shapeKeywords(27) = "CALLOUT_OVAL"
    shapeKeywords(28) = "CALLOUT_CLOUD"
    shapeKeywords(29) = "CALLOUT"

    Dim kk As Integer
    For kk = 0 To 29
        If InStr(upperLine, " " & shapeKeywords(kk) & " ") > 0 Or _
           InStr(upperLine, " " & shapeKeywords(kk) & vbCr) > 0 Or _
           right(upperLine, Len(shapeKeywords(kk))) = shapeKeywords(kk) Then
            shapeType = shapeKeywords(kk)
            Exit For
        End If
    Next kk

    If shapeType = "" Then
        IScr_Log "Line " & lineNum & ": ERROR - Unknown shape type"
        Exit Function
    End If

    Dim atPos As Integer
    atPos = InStr(upperLine, " AT ")
    If atPos = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - INSERT requires AT x, y"
        Exit Function
    End If

    Dim afterAt As String
    afterAt = Trim(Mid(line, atPos + 4))

    Dim commaPos As Integer
    commaPos = IScr_FindCommaOutsideParens(afterAt)
    If commaPos = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - AT requires x, y (comma separated)"
        Exit Function
    End If

    Dim xExpr As String
    Dim yAndRest As String
    xExpr = Trim(left(afterAt, commaPos - 1))
    yAndRest = Trim(Mid(afterAt, commaPos + 1))

    Dim widthPos As Integer
    widthPos = InStr(UCase(yAndRest), " WIDTH ")
    Dim yExpr As String
    If widthPos > 0 Then
        yExpr = Trim(left(yAndRest, widthPos - 1))
    Else
        yExpr = yAndRest
    End If

    Dim xVal As Single: xVal = CSng(IScr_ComputeNumber(xExpr))
    Dim yVal As Single: yVal = CSng(IScr_ComputeNumber(yExpr))

    Dim wVal As Single: wVal = CSng(IScr_EvalKeywordExpr(upperLine, line, "WIDTH"))
    Dim hVal As Single: hVal = CSng(IScr_EvalKeywordExpr(upperLine, line, "HEIGHT"))

    If wVal = -1 Then IScr_Log "Line " & lineNum & ": ERROR - INSERT requires WIDTH": Exit Function
    If hVal = -1 Then IScr_Log "Line " & lineNum & ": ERROR - INSERT requires HEIGHT": Exit Function

    Dim shapeName As String
    shapeName = IScr_ParseKeywordStringExpr(upperLine, line, "NAME")
    If shapeName = "" Then
        IScr_insertCounter = IScr_insertCounter + 1
        Select Case shapeType
            Case "RECTANGLE", "ROUNDEDRECTANGLE": shapeName = "script_rect_" & IScr_insertCounter
            Case "TEXTBOX":           shapeName = "script_text_" & IScr_insertCounter
            Case "OVAL":              shapeName = "script_oval_" & IScr_insertCounter
            Case "TRIANGLE", "RIGHTTRIANGLE": shapeName = "script_tri_" & IScr_insertCounter
            Case "DIAMOND":           shapeName = "script_diamond_" & IScr_insertCounter
            Case "PARALLELOGRAM":     shapeName = "script_para_" & IScr_insertCounter
            Case "TRAPEZOID":         shapeName = "script_trap_" & IScr_insertCounter
            Case "HEXAGON":           shapeName = "script_hex_" & IScr_insertCounter
            Case "PENTAGON", "PENTAGON_ARROW": shapeName = "script_pent_" & IScr_insertCounter
            Case "OCTAGON":           shapeName = "script_oct_" & IScr_insertCounter
            Case "ARROWRIGHT", "ARROWLEFT", "ARROWUP", "ARROWDOWN", "ARROWLEFTRIGHT": shapeName = "script_arrow_" & IScr_insertCounter
            Case "CHEVRON":           shapeName = "script_chev_" & IScr_insertCounter
            Case "CIRCULARRIGHTARROW": shapeName = "script_arrow_" & IScr_insertCounter
            Case "FLOWCHART_PROCESS", "FLOWCHART_DECISION", "FLOWCHART_TERMINATOR", _
                 "FLOWCHART_DATA", "FLOWCHART_DOCUMENT", "FLOWCHART_CONNECTOR": shapeName = "script_fc_" & IScr_insertCounter
            Case "CALLOUT_RECT", "CALLOUT_OVAL", "CALLOUT_CLOUD": shapeName = "script_callout_" & IScr_insertCounter
        End Select
    End If

    Dim shapeText As String
    shapeText = IScr_ParseKeywordStringExpr(upperLine, line, "TEXT")


    Dim oSlide As Slide
    Set oSlide = ActiveWindow.View.Slide

    If IScr_ShapeNameExists(oSlide, shapeName) Then
        IScr_Log "Line " & lineNum & ": ERROR - Shape """ & shapeName & """ already exists. Delete it first or use a different name."
        Exit Function
    End If

       Dim newShp As shape
    Select Case shapeType
        Case "RECTANGLE":         Set newShp = oSlide.shapes.AddShape(msoShapeRectangle, xVal, yVal, wVal, hVal)
        Case "TEXTBOX":           Set newShp = oSlide.shapes.AddTextbox(msoTextOrientationHorizontal, xVal, yVal, wVal, hVal)
        Case "OVAL":              Set newShp = oSlide.shapes.AddShape(msoShapeOval, xVal, yVal, wVal, hVal)
        Case "ROUNDEDRECTANGLE":  Set newShp = oSlide.shapes.AddShape(msoShapeRoundedRectangle, xVal, yVal, wVal, hVal)
        Case "TRIANGLE":          Set newShp = oSlide.shapes.AddShape(msoShapeIsoscelesTriangle, xVal, yVal, wVal, hVal)
        Case "RIGHTTRIANGLE":     Set newShp = oSlide.shapes.AddShape(msoShapeRightTriangle, xVal, yVal, wVal, hVal)
        Case "DIAMOND":           Set newShp = oSlide.shapes.AddShape(msoShapeDiamond, xVal, yVal, wVal, hVal)
        Case "PARALLELOGRAM":     Set newShp = oSlide.shapes.AddShape(msoShapeParallelogram, xVal, yVal, wVal, hVal)
        Case "TRAPEZOID":         Set newShp = oSlide.shapes.AddShape(msoShapeTrapezoid, xVal, yVal, wVal, hVal)
        Case "HEXAGON":           Set newShp = oSlide.shapes.AddShape(msoShapeHexagon, xVal, yVal, wVal, hVal)
        Case "PENTAGON":          Set newShp = oSlide.shapes.AddShape(msoShapeRegularPentagon, xVal, yVal, wVal, hVal)
        Case "OCTAGON":           Set newShp = oSlide.shapes.AddShape(msoShapeOctagon, xVal, yVal, wVal, hVal)
        Case "ARROWRIGHT":        Set newShp = oSlide.shapes.AddShape(msoShapeRightArrow, xVal, yVal, wVal, hVal)
        Case "ARROWLEFT":         Set newShp = oSlide.shapes.AddShape(msoShapeLeftArrow, xVal, yVal, wVal, hVal)
        Case "ARROWUP":           Set newShp = oSlide.shapes.AddShape(msoShapeUpArrow, xVal, yVal, wVal, hVal)
        Case "ARROWDOWN":         Set newShp = oSlide.shapes.AddShape(msoShapeDownArrow, xVal, yVal, wVal, hVal)
        Case "ARROWLEFTRIGHT":    Set newShp = oSlide.shapes.AddShape(msoShapeLeftRightArrow, xVal, yVal, wVal, hVal)
        Case "CHEVRON":           Set newShp = oSlide.shapes.AddShape(msoShapeChevron, xVal, yVal, wVal, hVal)
        Case "PENTAGON_ARROW":    Set newShp = oSlide.shapes.AddShape(msoShapePentagon, xVal, yVal, wVal, hVal)
        Case "CIRCULARRIGHTARROW": Set newShp = oSlide.shapes.AddShape(msoShapeCircularArrow, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_PROCESS":    Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartProcess, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_DECISION":   Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartDecision, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_TERMINATOR": Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartTerminator, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_DATA":       Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartData, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_DOCUMENT":   Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartDocument, xVal, yVal, wVal, hVal)
        Case "FLOWCHART_CONNECTOR":  Set newShp = oSlide.shapes.AddShape(msoShapeFlowchartConnector, xVal, yVal, wVal, hVal)
        Case "CALLOUT_RECT":  Set newShp = oSlide.shapes.AddShape(msoShapeRectangularCallout, xVal, yVal, wVal, hVal)
        Case "CALLOUT_OVAL":  Set newShp = oSlide.shapes.AddShape(msoShapeOvalCallout, xVal, yVal, wVal, hVal)
        Case "CALLOUT_CLOUD": Set newShp = oSlide.shapes.AddShape(msoShapeCloudCallout, xVal, yVal, wVal, hVal)
    End Select

    newShp.name = shapeName
    If shapeText <> "" Then
        If newShp.HasTextFrame Then newShp.TextFrame.textRange.text = shapeText
    End If

    Set IScr_ExecuteInsert = newShp
End Function

Public Sub IScr_ExecuteDelete(line As String, lineNum As Integer)
    Dim upperLine As String
    upperLine = UCase(Trim(line))

    Dim oSlide As Slide
    Set oSlide = ActiveWindow.View.Slide

    If upperLine = "DELETE SELECTION" Then
        On Error Resume Next
        Dim sel As Selection
        Set sel = ActiveWindow.Selection
        If sel.Type = ppSelectionShapes Then
            Dim count As Integer
            count = sel.ShapeRange.count
            sel.ShapeRange.Delete
            IScr_Log "Line " & lineNum & ": Deleted " & count & " selected shape(s)"
        Else
            IScr_Log "Line " & lineNum & ": WARNING - No shapes selected to delete"
        End If
        On Error GoTo 0
        Exit Sub
    End If

    If InStr(upperLine, "WHERE") = 0 Then
        IScr_Log "Line " & lineNum & ": ERROR - Expected WHERE or SELECTION after DELETE"
        Exit Sub
    End If

    Dim rawCriteria As String
    rawCriteria = Trim(Mid(line, InStr(upperLine, "WHERE") + 5))
    Dim criteria As String
    criteria = IScr_ResolveSelectCriteria(rawCriteria)

    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim shp As shape
    For Each shp In oSlide.shapes
        If IScr_ShapeMatchesCriteria(shp, criteria) Then toDelete.Add shp
    Next shp

    Dim deleted As Integer
    deleted = toDelete.count
    Dim i As Integer
    For i = 1 To toDelete.count
        toDelete(i).Delete
    Next i

    IScr_Log "Line " & lineNum & ": Deleted " & deleted & " shape(s)"
End Sub


Public Sub IScr_ExecuteRotate(line As String, shapes As Collection, lineNum As Integer)
    If shapes.count = 0 Then
        IScr_Log "Line " & lineNum & ": WARNING - ROTATE called but no shapes in working set"
        Exit Sub
    End If

    Dim upperLine As String
    upperLine = UCase(Trim(line))

    Dim relative As Boolean
    Dim angleExpr As String
    relative = False

    If left(upperLine, 9) = "ROTATE BY" Then
        relative = True
        angleExpr = Trim(Mid(line, 10))
    Else
        angleExpr = Trim(Mid(line, 7))
    End If

    Dim angle As Single
    angle = CSng(IScr_ComputeNumber(angleExpr))

    Dim shp As shape
    Dim count As Integer
    count = 0
    For Each shp In shapes
        On Error Resume Next
        If relative Then
            shp.rotation = shp.rotation + angle
        Else
            shp.rotation = angle
        End If
        If Err.Number = 0 Then count = count + 1
        On Error GoTo 0
    Next shp

    If relative Then
        IScr_Log "Line " & lineNum & ": ROTATE BY " & angle & "deg -> applied to " & count & " shape(s)"
    Else
        IScr_Log "Line " & lineNum & ": ROTATE " & angle & "deg -> applied to " & count & " shape(s)"
    End If
End Sub


Public Function IScr_ExecuteGroup(line As String, shapes As Collection, lineNum As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    If shapes.count < 2 Then
        IScr_Log "Line " & lineNum & ": ERROR - GROUP requires at least 2 shapes in working set"
        Set IScr_ExecuteGroup = shapes
        Exit Function
    End If

    On Error GoTo Failed

    Dim oSlide As Slide
    Set oSlide = ActiveWindow.View.Slide

    Dim names() As String
    ReDim names(1 To shapes.count)
    Dim i As Integer
    For i = 1 To shapes.count
        names(i) = shapes(i).name
    Next i

    Dim sr As ShapeRange
    Set sr = oSlide.shapes.Range(names)

    Dim grp As shape
    Set grp = sr.Group

    
    Dim upperLine As String
    upperLine = UCase(Trim(line))
    Dim grpName As String
    grpName = IScr_ParseKeywordStringExpr(upperLine, line, "NAME")
    If grpName <> "" Then grp.name = grpName

    result.Add grp
    IScr_SyncSelectionToPowerPoint result, lineNum
    IScr_Log "Line " & lineNum & ": Grouped " & shapes.count & " shape(s)" & IIf(grpName <> "", " as """ & grpName & """", "")

    Set IScr_ExecuteGroup = result
    Exit Function
Failed:
    IScr_Log "Line " & lineNum & ": ERROR - GROUP failed: " & Err.Description
    Set IScr_ExecuteGroup = shapes
End Function


Public Sub IScr_InvokeCommand(line As String, lineNum As Integer)
    Dim subName As String
    subName = Trim(Mid(line, 5))

    If subName = "" Then
        IScr_Log "Line " & lineNum & ": ERROR - CALL requires a sub name"
        Exit Sub
    End If

    On Error GoTo Failed
    Application.Run subName
    IScr_Log "Line " & lineNum & ": CALL " & subName & " - OK"
    Exit Sub
Failed:
    IScr_Log "Line " & lineNum & ": ERROR - CALL " & subName & " failed: " & Err.Description
End Sub

Public Sub IScr_Log(msg As String)
    If IScr_ScriptLog Is Nothing Then Set IScr_ScriptLog = New Collection
    IScr_ScriptLog.Add msg
End Sub

Public Function IScr_EvalKeywordExpr(upperLine As String, originalLine As String, keyword As String) As Double
    Dim pos As Integer
    pos = InStr(upperLine, " " & keyword & " ")
    If pos = 0 Then IScr_EvalKeywordExpr = -1: Exit Function

    Dim afterKeyword As String
    afterKeyword = Trim(Mid(originalLine, pos + Len(keyword) + 1))

    
    Dim keywords(4) As String
    keywords(0) = " NAME "
    keywords(1) = " TEXT "
    keywords(2) = " WIDTH "
    keywords(3) = " HEIGHT "
    keywords(4) = " AT "

    Dim endPos As Integer
    endPos = Len(afterKeyword) + 1
    Dim k As Integer
    For k = 0 To 4
        Dim kp As Integer
        kp = InStr(UCase(afterKeyword), keywords(k))
        If kp > 0 And kp < endPos Then endPos = kp
    Next k

    Dim exprStr As String
    exprStr = Trim(left(afterKeyword, endPos - 1))
    IScr_EvalKeywordExpr = IScr_ComputeNumber(exprStr)
End Function

Public Function IScr_ParseKeywordStringExpr(upperLine As String, originalLine As String, keyword As String) As String
    Dim pos As Integer
    pos = InStr(upperLine, " " & keyword & " ")
    If pos = 0 Then IScr_ParseKeywordStringExpr = "": Exit Function

    Dim afterKeyword As String
    afterKeyword = Trim(Mid(originalLine, pos + Len(keyword) + 1))

   
    Dim keywords(3) As String
    keywords(0) = " NAME "
    keywords(1) = " TEXT "
    keywords(2) = " WIDTH "
    keywords(3) = " HEIGHT "

    Dim endPos As Integer
    endPos = Len(afterKeyword) + 1
    Dim inQ As Boolean: inQ = False
    Dim i As Integer
    For i = 1 To Len(afterKeyword)
        If Mid(afterKeyword, i, 1) = """" Then inQ = Not inQ
        If Not inQ Then
            Dim k As Integer
            For k = 0 To 3
                If i + Len(keywords(k)) - 1 <= Len(afterKeyword) Then
                    If UCase(Mid(afterKeyword, i, Len(keywords(k)))) = keywords(k) Then
                        If i < endPos Then endPos = i
                    End If
                End If
            Next k
        End If
    Next i

    Dim exprStr As String
    exprStr = Trim(left(afterKeyword, endPos - 1))
    IScr_ParseKeywordStringExpr = IScr_ComputeText(exprStr)
End Function


Public Sub IScr_SyncSelectionToPowerPoint(shapes As Collection, lineNum As Integer)
    If shapes.count = 0 Then Exit Sub

    On Error GoTo Failed
    Dim oSlide As Slide
    Set oSlide = ActiveWindow.View.Slide

    Dim names() As String
    ReDim names(1 To shapes.count)
    Dim i As Integer
    For i = 1 To shapes.count
        names(i) = shapes(i).name
    Next i

    Dim sr As ShapeRange
    Set sr = oSlide.shapes.Range(names)
    sr.Select
    Exit Sub
Failed:
    IScr_Log "Line " & lineNum & ": WARNING - Could not sync selection to PowerPoint: " & Err.Description
End Sub


Public Function IScr_ExtractQuotedValue(s As String) As String
    Dim startQ As Integer
    startQ = InStr(s, """")
    If startQ = 0 Then
        Dim eqPos As Integer
        eqPos = InStr(s, "=")
        If eqPos > 0 Then
            IScr_ExtractQuotedValue = Trim(Mid(s, eqPos + 1))
        Else
            IScr_ExtractQuotedValue = Trim(s)
        End If
        Exit Function
    End If
    Dim endQ As Integer
    endQ = InStr(startQ + 1, s, """")
    If endQ = 0 Then
        IScr_ExtractQuotedValue = Mid(s, startQ + 1)
    Else
        IScr_ExtractQuotedValue = Mid(s, startQ + 1, endQ - startQ - 1)
    End If
End Function


Public Function IScr_ResolveSelectCriteria(criteria As String) As String

    Dim upperC As String
    upperC = UCase(Trim(criteria))

    Dim quotePos As Integer
    quotePos = InStr(criteria, """")

    If quotePos > 0 Then
        
        Dim eqPos As Integer
        eqPos = InStr(criteria, "=")
        Dim coPos As Integer
        coPos = InStr(upperC, "CONTAINS")
        Dim swPos As Integer
        swPos = InStr(upperC, "STARTSWITH")

        Dim splitAt As Integer
        If eqPos > 0 Then splitAt = eqPos
        If coPos > 0 Then splitAt = coPos + 8
        If swPos > 0 Then splitAt = swPos + 10

        If splitAt > 0 Then
            Dim keyword As String
            keyword = left(criteria, splitAt - 1)
            Dim valueExpr As String
            valueExpr = Trim(Mid(criteria, splitAt + 1))
            If left(keyword, Len(keyword)) = left(criteria, Len(keyword)) Then
                
                Dim resolved As String
                resolved = IScr_ComputeText(valueExpr)
                IScr_ResolveSelectCriteria = Trim(left(criteria, splitAt)) & " """ & resolved & """"
                Exit Function
            End If
        End If
    End If

    IScr_ResolveSelectCriteria = criteria
End Function

Public Function IScr_ShapeMatchesCriteria(shp As shape, criteria As String) As Boolean
    Dim upperCriteria As String
    upperCriteria = UCase(Trim(criteria))

    If left(upperCriteria, 7) = "NAME = " Then
        IScr_ShapeMatchesCriteria = (LCase(shp.name) = LCase(IScr_ExtractQuotedValue(criteria)))
        Exit Function
    End If

    If left(upperCriteria, 13) = "NAME CONTAINS" Then
        IScr_ShapeMatchesCriteria = (InStr(LCase(shp.name), LCase(IScr_ExtractQuotedValue(criteria))) > 0)
        Exit Function
    End If

    If left(upperCriteria, 15) = "NAME STARTSWITH" Then
        Dim prefix As String
        prefix = LCase(IScr_ExtractQuotedValue(criteria))
        IScr_ShapeMatchesCriteria = (left(LCase(shp.name), Len(prefix)) = prefix)
        Exit Function
    End If

    If left(upperCriteria, 7) = "TYPE = " Then
        IScr_ShapeMatchesCriteria = IScr_ShapeMatchesType(shp, Trim(Mid(upperCriteria, 8)))
        Exit Function
    End If

    IScr_ShapeMatchesCriteria = False
End Function

Public Function IScr_ShapeMatchesType(shp As shape, typeVal As String) As Boolean
    Select Case typeVal
        Case "RECTANGLE":        IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangle)
        Case "TEXTBOX":           IScr_ShapeMatchesType = (shp.Type = msoTextBox)
        Case "PICTURE":           IScr_ShapeMatchesType = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
        Case "TABLE":             IScr_ShapeMatchesType = shp.HasTable
        Case "LINE":              IScr_ShapeMatchesType = (shp.Type = msoLine)
        Case "OVAL":              IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOval)
        Case "ROUNDEDRECTANGLE":  IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRoundedRectangle)
        Case "TRIANGLE":          IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeIsoscelesTriangle)
        Case "RIGHTTRIANGLE":     IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRightTriangle)
        Case "DIAMOND":           IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeDiamond)
        Case "PARALLELOGRAM":     IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeParallelogram)
        Case "TRAPEZOID":         IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeTrapezoid)
        Case "HEXAGON":           IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeHexagon)
        Case "PENTAGON":          IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRegularPentagon)
        Case "OCTAGON":           IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOctagon)
        Case "ARROWRIGHT":        IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRightArrow)
        Case "ARROWLEFT":         IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeLeftArrow)
        Case "ARROWUP":           IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeUpArrow)
        Case "ARROWDOWN":         IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeDownArrow)
        Case "ARROWLEFTRIGHT":    IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeLeftRightArrow)
        Case "CHEVRON":           IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeChevron)
        Case "PENTAGON_ARROW":    IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapePentagon)
        Case "CIRCULARRIGHTARROW": IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeCircularArrow)
        Case "FLOWCHART_PROCESS":    IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartProcess)
        Case "FLOWCHART_DECISION":   IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartDecision)
        Case "FLOWCHART_TERMINATOR": IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartTerminator)
        Case "FLOWCHART_DATA":       IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartData)
        Case "FLOWCHART_DOCUMENT":   IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartDocument)
        Case "FLOWCHART_CONNECTOR":  IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartConnector)
        Case "CALLOUT_RECT":  IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangularCallout)
        Case "CALLOUT_OVAL":  IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOvalCallout)
        Case "CALLOUT_CLOUD": IScr_ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeCloudCallout)
        Case Else: IScr_ShapeMatchesType = False
    End Select
End Function




Public Function IScr_FindCommaOutsideParens(s As String) As Integer
    Dim depth As Integer: depth = 0
    Dim i As Integer
    For i = 1 To Len(s)
        Dim c As String: c = Mid(s, i, 1)
        If c = "(" Then depth = depth + 1
        If c = ")" Then depth = depth - 1
        If c = "," And depth = 0 Then
            IScr_FindCommaOutsideParens = i
            Exit Function
        End If
    Next i
    IScr_FindCommaOutsideParens = 0
End Function

Public Function IScr_ShapeNameExists(oSlide As Slide, shapeName As String) As Boolean
    Dim shp As shape
    For Each shp In oSlide.shapes
        If LCase(shp.name) = LCase(shapeName) Then
            IScr_ShapeNameExists = True
            Exit Function
        End If
    Next shp
    IScr_ShapeNameExists = False
End Function




