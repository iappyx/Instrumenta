Attribute VB_Name = "ModuleInstrumentaScript"
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

Public ScriptLog As Collection

Private varNames() As String
Private varValues() As Double
Private varCount As Integer

Private breakFlag As Boolean
Private insertCounter As Integer
Private exprTokens() As String
Private exprPos As Integer

Public Sub ShowScriptEditor()
    ScriptEditorForm.Show vbModeless
End Sub


Public Sub RunScript(scriptText As String)
    Set ScriptLog = New Collection
    insertCounter = 0
    breakFlag = False
    varCount = 0
    ReDim varNames(0)
    ReDim varValues(0)

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

    ExecuteBlock lines, 0, UBound(lines), lineNum, selectedShapes

    Log "---"
    Log "Done."
End Sub

Private Function ExecuteBlock(lines() As String, startIdx As Integer, endIdx As Integer, _
                               ByRef lineNum As Integer, ByRef selectedShapes As Collection) As Integer
    Dim i As Integer
    i = startIdx

    Do While i <= endIdx
        If breakFlag Then Exit Do

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
            Set selectedShapes = ExecuteSelect(line, lineNum)
            SyncSelectionToPowerPoint selectedShapes, lineNum
            Log "Line " & lineNum & ": Selected " & selectedShapes.count & " shape(s)"

        
        ElseIf upperLine = "USE SELECTION" Then
            Set selectedShapes = ExecuteUseSelection(lineNum)
            Log "Line " & lineNum & ": Using PowerPoint selection - " & selectedShapes.count & " shape(s)"

        
        ElseIf left(upperLine, 6) = "INSERT" Then
            Dim newShape As shape
            Set newShape = ExecuteInsert(line, lineNum)
            If Not newShape Is Nothing Then
                Set selectedShapes = New Collection
                selectedShapes.Add newShape
                SyncSelectionToPowerPoint selectedShapes, lineNum
                Log "Line " & lineNum & ": Inserted """ & newShape.name & """ - now working set"
            Else
                Set selectedShapes = New Collection
            End If

        
        ElseIf left(upperLine, 6) = "DELETE" Then
            ExecuteDelete line, lineNum
            Set selectedShapes = New Collection

        
        ElseIf left(upperLine, 3) = "SET" Then
            If selectedShapes.count = 0 Then
                Log "Line " & lineNum & ": WARNING - SET called but no shapes in working set"
            Else
                ExecuteSet selectedShapes, line, lineNum
            End If

        
        ElseIf left(upperLine, 4) = "CALL" Then
            ExecuteCall line, lineNum

        ElseIf left(upperLine, 5) = "GROUP" Then
            Set selectedShapes = ExecuteGroup(line, selectedShapes, lineNum)

        ElseIf left(upperLine, 6) = "ROTATE" Then
            ExecuteRotate line, selectedShapes, lineNum

        ElseIf left(upperLine, 6) = "REPEAT" Then
           
            Dim repeatEnd As Integer
            repeatEnd = FindBlockEnd(lines, i, "REPEAT", "END REPEAT")
            If repeatEnd = -1 Then
                Log "Line " & lineNum & ": ERROR - No matching END REPEAT found"
                i = endIdx + 1
                GoTo NextIteration
            End If
            ExecuteRepeat lines, i, repeatEnd, lineNum, selectedShapes
            i = repeatEnd + 1
            GoTo NextIteration

        ElseIf left(upperLine, 2) = "IF" And left(upperLine, 5) <> "IF i " Or _
               (left(upperLine, 2) = "IF" And IsIfCommand(upperLine)) Then
            Dim ifEnd As Integer
            ifEnd = FindBlockEnd(lines, i, "IF", "END IF")
            If ifEnd = -1 Then
                Log "Line " & lineNum & ": ERROR - No matching END IF found"
                i = endIdx + 1
                GoTo NextIteration
            End If
            ExecuteIf lines, i, ifEnd, lineNum, selectedShapes
            i = ifEnd + 1
            GoTo NextIteration

        ElseIf upperLine = "BREAK" Then
            Log "Line " & lineNum & ": BREAK"
            breakFlag = True
            Exit Do


        ElseIf upperLine = "END REPEAT" Or upperLine = "END IF" Or _
               left(upperLine, 4) = "ELSE" Then

            i = i + 1
            GoTo NextIteration

        Else
            Log "Line " & lineNum & ": ERROR - Unknown command: " & Trim(line)
        End If

        i = i + 1
NextIteration:
    Loop

    ExecuteBlock = i - 1
End Function

Private Function IsIfCommand(upperLine As String) As Boolean
    IsIfCommand = (left(upperLine, 3) = "IF ")
End Function


Private Sub ExecuteRepeat(lines() As String, repeatLine As Integer, endLine As Integer, _
                           ByRef lineNum As Integer, ByRef selectedShapes As Collection)
    Dim headerLine As String
    headerLine = UCase(Trim(lines(repeatLine)))


    Dim parts As String
    parts = Trim(Mid(headerLine, 7))

    Dim countVal As Double
    Dim asPos As Integer
    asPos = InStr(parts, " AS ")
    If asPos = 0 Then
        Log "Line " & (repeatLine + 1) & ": ERROR - REPEAT requires AS <variable>"
        Exit Sub
    End If
    countVal = EvalNumericExpr(Trim(left(parts, asPos - 1)))

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
            startVal = EvalNumericExpr(Trim(left(afterFrom, stepPos2 - 1)))
        Else
            startVal = EvalNumericExpr(Trim(afterFrom))
        End If
    End If

    Dim stepVal As Double
    stepVal = 1
    Dim stepPos As Integer
    stepPos = InStr(parts, " STEP ")
    If stepPos > 0 Then
        stepVal = EvalNumericExpr(Trim(Mid(parts, stepPos + 6)))
    End If

    Dim bodyStart As Integer
    Dim bodyEnd As Integer
    bodyStart = repeatLine + 1
    bodyEnd = endLine - 1

    Dim loopVal As Double
    Dim iteration As Long
    For iteration = 0 To countVal - 1
        loopVal = startVal + iteration * stepVal
        SetVar varName, loopVal
        breakFlag = False
        ExecuteBlock lines, bodyStart, bodyEnd, lineNum, selectedShapes
        If breakFlag Then
            breakFlag = False
            Exit For
        End If
    Next iteration
End Sub


Private Sub ExecuteIf(lines() As String, ifLine As Integer, endIfLine As Integer, _
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
            condResult = EvalCondition(branchConds(b))
        End If

        If condResult Then
            ExecuteBlock lines, branchStarts(b), branchEnds(b), lineNum, selectedShapes
            Exit For
        End If
    Next b
End Sub


Private Function FindBlockEnd(lines() As String, startIdx As Integer, _
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
                FindBlockEnd = i
                Exit Function
            End If
            depth = depth - 1
        End If
    Next i

    FindBlockEnd = -1
End Function


Private Sub SetVar(varName As String, value As Double)
    Dim i As Integer
    For i = 0 To varCount - 1
        If varNames(i) = varName Then
            varValues(i) = value
            Exit Sub
        End If
    Next i
    ReDim Preserve varNames(varCount)
    ReDim Preserve varValues(varCount)
    varNames(varCount) = varName
    varValues(varCount) = value
    varCount = varCount + 1
End Sub

Private Function GetVar(varName As String) As Double
    Dim i As Integer
    For i = 0 To varCount - 1
        If varNames(i) = LCase(varName) Then
            GetVar = varValues(i)
            Exit Function
        End If
    Next i
    GetVar = 0
End Function

Private Function VarExists(varName As String) As Boolean
    Dim i As Integer
    For i = 0 To varCount - 1
        If varNames(i) = LCase(varName) Then
            VarExists = True
            Exit Function
        End If
    Next i
    VarExists = False
End Function

Public Function EvalNumericExpr(expr As String) As Double

    Dim substituted As String
    substituted = SubstituteVars(expr)

    exprTokens = TokenizeExpr(substituted)
    exprPos = 0

    On Error GoTo Failed
    EvalNumericExpr = ParseExprAddSub()
    Exit Function
Failed:
    EvalNumericExpr = 0
End Function

Private Function SubstituteVars(expr As String) As String

    Dim result As String
    result = expr
    Dim i As Integer
    For i = 0 To varCount - 1
        result = ReplaceWholeWord(result, varNames(i), CStr(varValues(i)))
    Next i
    SubstituteVars = result
End Function

Private Function ReplaceWholeWord(s As String, word As String, replacement As String) As String
    Dim result As String
    Dim i As Integer
    i = 1
    result = ""
    Dim sLen As Integer
    sLen = Len(s)
    Dim wLen As Integer
    wLen = Len(word)

    Do While i <= sLen
        If i + wLen - 1 <= sLen Then
            If LCase(Mid(s, i, wLen)) = LCase(word) Then
             
                Dim beforeOk As Boolean
                Dim afterOk As Boolean

                If i = 1 Then
                    beforeOk = True
                Else
                    beforeOk = Not IsAlphaNum(Mid(s, i - 1, 1))
                End If

                If i + wLen > sLen Then
                    afterOk = True
                Else
                    afterOk = Not IsAlphaNum(Mid(s, i + wLen, 1))
                End If

                If beforeOk And afterOk Then
                    result = result & replacement
                    i = i + wLen
                    GoTo ContinueLoop
                End If
            End If
        End If
        result = result & Mid(s, i, 1)
        i = i + 1
ContinueLoop:
    Loop
    ReplaceWholeWord = result
End Function

Private Function IsAlphaNum(c As String) As Boolean
    IsAlphaNum = (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or _
                 (c >= "0" And c <= "9") Or c = "_"
End Function

Private Function TokenizeExpr(expr As String) As String()
    Dim tokens() As String
    Dim count As Integer
    count = 0
    ReDim tokens(0)

    Dim i As Integer
    i = 1
    Dim s As String
    s = Trim(expr)
    Dim sLen As Integer
    sLen = Len(s)

    Do While i <= sLen
        Dim c As String
        c = Mid(s, i, 1)

        If c = " " Then
            i = i + 1

        ElseIf c = "(" Or c = ")" Or c = "*" Or c = "/" Then
            ReDim Preserve tokens(count)
            tokens(count) = c
            count = count + 1
            i = i + 1

        ElseIf c = "+" Or c = "-" Then
            ReDim Preserve tokens(count)
            tokens(count) = c
            count = count + 1
            i = i + 1

        ElseIf (c >= "0" And c <= "9") Or c = "." Then
            
            Dim numStr As String
            numStr = ""
            Do While i <= sLen
                Dim nc As String
                nc = Mid(s, i, 1)
                If (nc >= "0" And nc <= "9") Or nc = "." Then
                    numStr = numStr & nc
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            ReDim Preserve tokens(count)
            tokens(count) = numStr
            count = count + 1

        Else
            
            i = i + 1
        End If
    Loop

    TokenizeExpr = tokens
End Function

Private Function ParseExprAddSub() As Double
    Dim left As Double
    left = ParseExprMulDiv()

    Do While exprPos <= UBound(exprTokens)
        Dim op As String
        op = exprTokens(exprPos)
        If op = "+" Or op = "-" Then
            exprPos = exprPos + 1
            Dim right As Double
            right = ParseExprMulDiv()
            If op = "+" Then left = left + right Else left = left - right
        Else
            Exit Do
        End If
    Loop

    ParseExprAddSub = left
End Function

Private Function ParseExprMulDiv() As Double
    Dim left As Double
    left = ParseExprUnary()

    Do While exprPos <= UBound(exprTokens)
        Dim op As String
        op = exprTokens(exprPos)
        If op = "*" Or op = "/" Then
            exprPos = exprPos + 1
            Dim right As Double
            right = ParseExprUnary()
            If op = "*" Then
                left = left * right
            Else
                If right <> 0 Then left = left / right Else left = 0
            End If
        Else
            Exit Do
        End If
    Loop

    ParseExprMulDiv = left
End Function

Private Function ParseExprUnary() As Double
    If exprPos <= UBound(exprTokens) Then
        If exprTokens(exprPos) = "-" Then
            exprPos = exprPos + 1
            ParseExprUnary = -ParseExprPrimary()
            Exit Function
        ElseIf exprTokens(exprPos) = "+" Then
            exprPos = exprPos + 1
        End If
    End If
    ParseExprUnary = ParseExprPrimary()
End Function

Private Function ParseExprPrimary() As Double
    If exprPos > UBound(exprTokens) Then
        ParseExprPrimary = 0
        Exit Function
    End If

    Dim tok As String
    tok = exprTokens(exprPos)

    If tok = "(" Then
        exprPos = exprPos + 1
        Dim val As Double
        val = ParseExprAddSub()
        If exprPos <= UBound(exprTokens) And exprTokens(exprPos) = ")" Then
            exprPos = exprPos + 1
        End If
        ParseExprPrimary = val
    Else
        exprPos = exprPos + 1
        ParseExprPrimary = CDbl(tok)
    End If
End Function


Public Function EvalStringExpr(expr As String) As String
    
    Dim parts() As String
    parts = SplitStringExpr(expr)

    Dim result As String
    result = ""
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim part As String
        part = Trim(parts(i))
        If left(part, 1) = """" Then
            
            result = result & Mid(part, 2, Len(part) - 2)
        ElseIf IsNumericExpr(part) Then
            
            Dim numVal As Double
            numVal = EvalNumericExpr(part)
            
            If numVal = Int(numVal) Then
                result = result & CStr(CLng(numVal))
            Else
                result = result & CStr(numVal)
            End If
        Else
            result = result & part
        End If
    Next i

    EvalStringExpr = result
End Function

Private Function IsNumericExpr(s As String) As Boolean
    IsNumericExpr = (InStr(s, """") = 0)
End Function

Private Function SplitStringExpr(expr As String) As String()
    Dim parts() As String
    Dim count As Integer
    count = 0
    ReDim parts(0)

    Dim i As Integer
    Dim inQuote As Boolean
    Dim current As String
    inQuote = False
    current = ""

    For i = 1 To Len(expr)
        Dim c As String
        c = Mid(expr, i, 1)
        If c = """" Then
            inQuote = Not inQuote
            current = current & c
        ElseIf c = "+" And Not inQuote Then
            ReDim Preserve parts(count)
            parts(count) = current
            count = count + 1
            current = ""
        Else
            current = current & c
        End If
    Next i

    ReDim Preserve parts(count)
    parts(count) = current

    SplitStringExpr = parts
End Function


Public Function EvalCondition(cond As String) As Boolean
    On Error GoTo Failed
    EvalCondition = ParseCondOr(Trim(cond))
    Exit Function
Failed:
    EvalCondition = False
End Function

Private Function ParseCondOr(cond As String) As Boolean
   
    Dim parts() As String
    parts = SplitCondOn(cond, " OR ")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If ParseCondAnd(Trim(parts(i))) Then
            ParseCondOr = True
            Exit Function
        End If
    Next i
    ParseCondOr = False
End Function

Private Function ParseCondAnd(cond As String) As Boolean
    
    Dim parts() As String
    parts = SplitCondOn(cond, " AND ")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If Not ParseCondNot(Trim(parts(i))) Then
            ParseCondAnd = False
            Exit Function
        End If
    Next i
    ParseCondAnd = True
End Function

Private Function ParseCondNot(cond As String) As Boolean
    Dim c As String
    c = Trim(cond)
    If left(UCase(c), 4) = "NOT " Then
        ParseCondNot = Not ParseCondAtom(Trim(Mid(c, 5)))
    Else
        ParseCondNot = ParseCondAtom(c)
    End If
End Function

Private Function ParseCondAtom(cond As String) As Boolean
    Dim c As String
    c = Trim(cond)

    If left(c, 1) = "(" And right(c, 1) = ")" Then
        ParseCondAtom = ParseCondOr(Trim(Mid(c, 2, Len(c) - 2)))
        Exit Function
    End If

    Dim ops(5) As String
    ops(0) = ">="
    ops(1) = "<="
    ops(2) = "<>"
    ops(3) = ">"
    ops(4) = "<"
    ops(5) = "="

    Dim opIdx As Integer
    For opIdx = 0 To 5
        Dim opStr As String
        opStr = ops(opIdx)
        Dim opPos As Integer
        opPos = InStr(c, opStr)
        If opPos > 0 Then
            Dim leftExpr As String
            Dim rightExpr As String
            leftExpr = Trim(left(c, opPos - 1))
            rightExpr = Trim(Mid(c, opPos + Len(opStr)))

            Dim leftVal As Double
            Dim rightVal As Double
            leftVal = EvalNumericExpr(leftExpr)
            rightVal = EvalNumericExpr(rightExpr)

            Select Case opStr
                Case ">=": ParseCondAtom = (leftVal >= rightVal)
                Case "<=": ParseCondAtom = (leftVal <= rightVal)
                Case "<>": ParseCondAtom = (leftVal <> rightVal)
                Case ">": ParseCondAtom = (leftVal > rightVal)
                Case "<": ParseCondAtom = (leftVal < rightVal)
                Case "=": ParseCondAtom = (leftVal = rightVal)
            End Select
            Exit Function
        End If
    Next opIdx

    If UCase(c) = "TRUE" Then ParseCondAtom = True
    If UCase(c) = "FALSE" Then ParseCondAtom = False
End Function

Private Function SplitCondOn(cond As String, separator As String) As String()
   
    Dim parts() As String
    Dim count As Integer
    count = 0
    ReDim parts(0)

    Dim i As Integer
    Dim depth As Integer
    depth = 0
    Dim current As String
    current = ""
    Dim sepLen As Integer
    sepLen = Len(separator)
    Dim condLen As Integer
    condLen = Len(cond)
    Dim upperCond As String
    upperCond = UCase(cond)

    i = 1
    Do While i <= condLen
        Dim c As String
        c = Mid(cond, i, 1)

        If c = "(" Then
            depth = depth + 1
            current = current & c
            i = i + 1
        ElseIf c = ")" Then
            depth = depth - 1
            current = current & c
            i = i + 1
        ElseIf depth = 0 And i + sepLen - 1 <= condLen Then
            If UCase(Mid(cond, i, sepLen)) = UCase(separator) Then
                ReDim Preserve parts(count)
                parts(count) = current
                count = count + 1
                current = ""
                i = i + sepLen
            Else
                current = current & c
                i = i + 1
            End If
        Else
            current = current & c
            i = i + 1
        End If
    Loop

    ReDim Preserve parts(count)
    parts(count) = current

    SplitCondOn = parts
End Function


Private Function ExecuteUseSelection(lineNum As Integer) As Collection
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
        Log "Line " & lineNum & ": WARNING - No shapes in current PowerPoint selection"
    End If

    Set ExecuteUseSelection = result
    Exit Function
Failed:
    Log "Line " & lineNum & ": ERROR - Could not read PowerPoint selection"
    Set ExecuteUseSelection = result
End Function


Private Function ExecuteSelect(line As String, lineNum As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim upperLine As String
    upperLine = UCase(Trim(line))

    Dim oSlide As Slide
    On Error Resume Next
    Set oSlide = ActiveWindow.View.Slide
    On Error GoTo 0

    If oSlide Is Nothing Then
        Log "Line " & lineNum & ": ERROR - No active slide"
        Set ExecuteSelect = result
        Exit Function
    End If

    If upperLine = "SELECT ALL" Then
        Dim shp As shape
        For Each shp In oSlide.shapes
            result.Add shp
        Next shp
        Set ExecuteSelect = result
        Exit Function
    End If

    If InStr(upperLine, "WHERE") = 0 Then
        Log "Line " & lineNum & ": ERROR - Expected ALL or WHERE after SELECT"
        Set ExecuteSelect = result
        Exit Function
    End If

    Dim rawCriteria As String
    rawCriteria = Trim(Mid(line, InStr(upperLine, "WHERE") + 5))
    Dim criteria As String
    criteria = ResolveSelectCriteria(rawCriteria)

    For Each shp In oSlide.shapes
        If ShapeMatchesCriteria(shp, criteria) Then
            result.Add shp
        End If
    Next shp

    Set ExecuteSelect = result
End Function

Private Function ResolveSelectCriteria(criteria As String) As String

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
                resolved = EvalStringExpr(valueExpr)
                ResolveSelectCriteria = Trim(left(criteria, splitAt)) & " """ & resolved & """"
                Exit Function
            End If
        End If
    End If

    ResolveSelectCriteria = criteria
End Function

Private Function ShapeMatchesCriteria(shp As shape, criteria As String) As Boolean
    Dim upperCriteria As String
    upperCriteria = UCase(Trim(criteria))

    If left(upperCriteria, 7) = "NAME = " Then
        ShapeMatchesCriteria = (LCase(shp.name) = LCase(ExtractQuotedValue(criteria)))
        Exit Function
    End If

    If left(upperCriteria, 13) = "NAME CONTAINS" Then
        ShapeMatchesCriteria = (InStr(LCase(shp.name), LCase(ExtractQuotedValue(criteria))) > 0)
        Exit Function
    End If

    If left(upperCriteria, 15) = "NAME STARTSWITH" Then
        Dim prefix As String
        prefix = LCase(ExtractQuotedValue(criteria))
        ShapeMatchesCriteria = (left(LCase(shp.name), Len(prefix)) = prefix)
        Exit Function
    End If

    If left(upperCriteria, 7) = "TYPE = " Then
        ShapeMatchesCriteria = ShapeMatchesType(shp, Trim(Mid(upperCriteria, 8)))
        Exit Function
    End If

    ShapeMatchesCriteria = False
End Function

Private Function ShapeMatchesType(shp As shape, typeVal As String) As Boolean
    Select Case typeVal
        Case "RECTANGLE":        ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangle)
        Case "TEXTBOX":           ShapeMatchesType = (shp.Type = msoTextBox)
        Case "PICTURE":           ShapeMatchesType = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
        Case "TABLE":             ShapeMatchesType = shp.HasTable
        Case "LINE":              ShapeMatchesType = (shp.Type = msoLine)
        Case "OVAL":              ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOval)
        Case "ROUNDEDRECTANGLE":  ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRoundedRectangle)
        Case "TRIANGLE":          ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeIsoscelesTriangle)
        Case "RIGHTTRIANGLE":     ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRightTriangle)
        Case "DIAMOND":           ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeDiamond)
        Case "PARALLELOGRAM":     ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeParallelogram)
        Case "TRAPEZOID":         ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeTrapezoid)
        Case "HEXAGON":           ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeHexagon)
        Case "PENTAGON":          ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRegularPentagon)
        Case "OCTAGON":           ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOctagon)
        Case "ARROWRIGHT":        ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRightArrow)
        Case "ARROWLEFT":         ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeLeftArrow)
        Case "ARROWUP":           ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeUpArrow)
        Case "ARROWDOWN":         ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeDownArrow)
        Case "ARROWLEFTRIGHT":    ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeLeftRightArrow)
        Case "CHEVRON":           ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeChevron)
        Case "PENTAGON_ARROW":    ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapePentagon)
        Case "CIRCULARRIGHTARROW": ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeCircularArrow)
        Case "FLOWCHART_PROCESS":    ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartProcess)
        Case "FLOWCHART_DECISION":   ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartDecision)
        Case "FLOWCHART_TERMINATOR": ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartTerminator)
        Case "FLOWCHART_DATA":       ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartData)
        Case "FLOWCHART_DOCUMENT":   ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartDocument)
        Case "FLOWCHART_CONNECTOR":  ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeFlowchartConnector)
        Case "CALLOUT_RECT":  ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRectangularCallout)
        Case "CALLOUT_OVAL":  ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOvalCallout)
        Case "CALLOUT_CLOUD": ShapeMatchesType = (shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeCloudCallout)
        Case Else: ShapeMatchesType = False
    End Select
End Function


Private Function ExecuteInsert(line As String, lineNum As Integer) As shape
    Set ExecuteInsert = Nothing

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
        Log "Line " & lineNum & ": ERROR - Unknown shape type"
        Exit Function
    End If

    Dim atPos As Integer
    atPos = InStr(upperLine, " AT ")
    If atPos = 0 Then
        Log "Line " & lineNum & ": ERROR - INSERT requires AT x, y"
        Exit Function
    End If

    Dim afterAt As String
    afterAt = Trim(Mid(line, atPos + 4))

    Dim commaPos As Integer
    commaPos = FindCommaOutsideParens(afterAt)
    If commaPos = 0 Then
        Log "Line " & lineNum & ": ERROR - AT requires x, y (comma separated)"
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

    Dim xVal As Single: xVal = CSng(EvalNumericExpr(xExpr))
    Dim yVal As Single: yVal = CSng(EvalNumericExpr(yExpr))

    Dim wVal As Single: wVal = CSng(EvalKeywordExpr(upperLine, line, "WIDTH"))
    Dim hVal As Single: hVal = CSng(EvalKeywordExpr(upperLine, line, "HEIGHT"))

    If wVal = -1 Then Log "Line " & lineNum & ": ERROR - INSERT requires WIDTH": Exit Function
    If hVal = -1 Then Log "Line " & lineNum & ": ERROR - INSERT requires HEIGHT": Exit Function

    Dim shapeName As String
    shapeName = ParseKeywordStringExpr(upperLine, line, "NAME")
    If shapeName = "" Then
        insertCounter = insertCounter + 1
        Select Case shapeType
            Case "RECTANGLE", "ROUNDEDRECTANGLE": shapeName = "script_rect_" & insertCounter
            Case "TEXTBOX":           shapeName = "script_text_" & insertCounter
            Case "OVAL":              shapeName = "script_oval_" & insertCounter
            Case "TRIANGLE", "RIGHTTRIANGLE": shapeName = "script_tri_" & insertCounter
            Case "DIAMOND":           shapeName = "script_diamond_" & insertCounter
            Case "PARALLELOGRAM":     shapeName = "script_para_" & insertCounter
            Case "TRAPEZOID":         shapeName = "script_trap_" & insertCounter
            Case "HEXAGON":           shapeName = "script_hex_" & insertCounter
            Case "PENTAGON", "PENTAGON_ARROW": shapeName = "script_pent_" & insertCounter
            Case "OCTAGON":           shapeName = "script_oct_" & insertCounter
            Case "ARROWRIGHT", "ARROWLEFT", "ARROWUP", "ARROWDOWN", "ARROWLEFTRIGHT": shapeName = "script_arrow_" & insertCounter
            Case "CHEVRON":           shapeName = "script_chev_" & insertCounter
            Case "CIRCULARRIGHTARROW": shapeName = "script_arrow_" & insertCounter
            Case "FLOWCHART_PROCESS", "FLOWCHART_DECISION", "FLOWCHART_TERMINATOR", _
                 "FLOWCHART_DATA", "FLOWCHART_DOCUMENT", "FLOWCHART_CONNECTOR": shapeName = "script_fc_" & insertCounter
            Case "CALLOUT_RECT", "CALLOUT_OVAL", "CALLOUT_CLOUD": shapeName = "script_callout_" & insertCounter
        End Select
    End If

    Dim shapeText As String
    shapeText = ParseKeywordStringExpr(upperLine, line, "TEXT")


    Dim oSlide As Slide
    Set oSlide = ActiveWindow.View.Slide

    If ShapeNameExists(oSlide, shapeName) Then
        Log "Line " & lineNum & ": ERROR - Shape """ & shapeName & """ already exists. Delete it first or use a different name."
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

    Set ExecuteInsert = newShp
End Function

Private Function FindCommaOutsideParens(s As String) As Integer
    Dim depth As Integer: depth = 0
    Dim i As Integer
    For i = 1 To Len(s)
        Dim c As String: c = Mid(s, i, 1)
        If c = "(" Then depth = depth + 1
        If c = ")" Then depth = depth - 1
        If c = "," And depth = 0 Then
            FindCommaOutsideParens = i
            Exit Function
        End If
    Next i
    FindCommaOutsideParens = 0
End Function

Private Function ShapeNameExists(oSlide As Slide, shapeName As String) As Boolean
    Dim shp As shape
    For Each shp In oSlide.shapes
        If LCase(shp.name) = LCase(shapeName) Then
            ShapeNameExists = True
            Exit Function
        End If
    Next shp
    ShapeNameExists = False
End Function


Private Sub ExecuteDelete(line As String, lineNum As Integer)
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
            Log "Line " & lineNum & ": Deleted " & count & " selected shape(s)"
        Else
            Log "Line " & lineNum & ": WARNING - No shapes selected to delete"
        End If
        On Error GoTo 0
        Exit Sub
    End If

    If InStr(upperLine, "WHERE") = 0 Then
        Log "Line " & lineNum & ": ERROR - Expected WHERE or SELECTION after DELETE"
        Exit Sub
    End If

    Dim rawCriteria As String
    rawCriteria = Trim(Mid(line, InStr(upperLine, "WHERE") + 5))
    Dim criteria As String
    criteria = ResolveSelectCriteria(rawCriteria)

    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim shp As shape
    For Each shp In oSlide.shapes
        If ShapeMatchesCriteria(shp, criteria) Then toDelete.Add shp
    Next shp

    Dim deleted As Integer
    deleted = toDelete.count
    Dim i As Integer
    For i = 1 To toDelete.count
        toDelete(i).Delete
    Next i

    Log "Line " & lineNum & ": Deleted " & deleted & " shape(s)"
End Sub


Private Sub ExecuteRotate(line As String, shapes As Collection, lineNum As Integer)
    If shapes.count = 0 Then
        Log "Line " & lineNum & ": WARNING - ROTATE called but no shapes in working set"
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
    angle = CSng(EvalNumericExpr(angleExpr))

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
        Log "Line " & lineNum & ": ROTATE BY " & angle & "deg -> applied to " & count & " shape(s)"
    Else
        Log "Line " & lineNum & ": ROTATE " & angle & "deg -> applied to " & count & " shape(s)"
    End If
End Sub


Private Function ExecuteGroup(line As String, shapes As Collection, lineNum As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    If shapes.count < 2 Then
        Log "Line " & lineNum & ": ERROR - GROUP requires at least 2 shapes in working set"
        Set ExecuteGroup = shapes
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
    grpName = ParseKeywordStringExpr(upperLine, line, "NAME")
    If grpName <> "" Then grp.name = grpName

    result.Add grp
    SyncSelectionToPowerPoint result, lineNum
    Log "Line " & lineNum & ": Grouped " & shapes.count & " shape(s)" & IIf(grpName <> "", " as """ & grpName & """", "")

    Set ExecuteGroup = result
    Exit Function
Failed:
    Log "Line " & lineNum & ": ERROR - GROUP failed: " & Err.Description
    Set ExecuteGroup = shapes
End Function


Private Sub ExecuteCall(line As String, lineNum As Integer)
    Dim subName As String
    subName = Trim(Mid(line, 5))

    If subName = "" Then
        Log "Line " & lineNum & ": ERROR - CALL requires a sub name"
        Exit Sub
    End If

    On Error GoTo Failed
    Application.Run subName
    Log "Line " & lineNum & ": CALL " & subName & " - OK"
    Exit Sub
Failed:
    Log "Line " & lineNum & ": ERROR - CALL " & subName & " failed: " & Err.Description
End Sub


Private Sub ExecuteSet(shapes As Collection, line As String, lineNum As Integer)
    Dim rest As String
    rest = Trim(Mid(line, 4))

    Dim eqPos As Integer
    eqPos = InStr(rest, "=")
    If eqPos = 0 Then
        Log "Line " & lineNum & ": ERROR - SET requires = sign"
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
        If ApplyProperty(shp, prop, valueExpr, lineNum) Then successCount = successCount + 1
    Next shp

    Log "Line " & lineNum & ": SET " & prop & " = " & valueExpr & " -> applied to " & successCount & " shape(s)"
End Sub

Private Function ApplyProperty(shp As shape, prop As String, valueExpr As String, lineNum As Integer) As Boolean
    ApplyProperty = True
    On Error GoTo Failed

    Dim numVal As Double
    Dim strVal As String

    Select Case prop
        Case "font.size", "width", "height", "position.x", "position.y", "opacity", "border.width"
            numVal = EvalNumericExpr(valueExpr)
        Case "font.name", "name"
            strVal = EvalStringExpr(valueExpr)
        Case "font.color", "fill.color", "border.color"
            strVal = Trim(valueExpr)
        Case "font.bold", "font.italic", "font.underline", "fill.transparent", "border.visible", "border.style"
            strVal = UCase(Trim(valueExpr))
    End Select

    Select Case prop
        Case "font.size":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Size = numVal
        Case "font.bold":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Bold = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.italic":     If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Italic = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.underline":  If shp.HasTextFrame Then shp.TextFrame.textRange.Font.Underline = (strVal = "TRUE" Or strVal = "1" Or strVal = "YES")
        Case "font.color":      If shp.HasTextFrame Then shp.TextFrame.textRange.Font.color.RGB = HexToRGB(strVal)
        Case "font.name":       If shp.HasTextFrame Then shp.TextFrame.textRange.Font.name = strVal
        Case "fill.color":      shp.Fill.Solid: shp.Fill.ForeColor.RGB = HexToRGB(strVal)
        Case "fill.transparent": If (strVal = "TRUE" Or strVal = "1") Then shp.Fill.visible = msoFalse Else shp.Fill.visible = msoTrue
        Case "width":           shp.width = CSng(numVal)
        Case "height":          shp.height = CSng(numVal)
        Case "position.x":      shp.left = CSng(numVal)
        Case "position.y":      shp.Top = CSng(numVal)
        Case "opacity":         shp.Fill.Transparency = 1 - (CSng(numVal) / 100)
        Case "name":            shp.name = strVal
        Case "border.color":
            shp.line.ForeColor.RGB = HexToRGB(strVal)
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
            Log "  WARNING - Unknown property: " & prop
            ApplyProperty = False
    End Select
    Exit Function
Failed:
    Log "  ERROR - Could not set " & prop & " on """ & shp.name & """: " & Err.Description
    ApplyProperty = False
End Function


Private Function EvalKeywordExpr(upperLine As String, originalLine As String, keyword As String) As Double
    Dim pos As Integer
    pos = InStr(upperLine, " " & keyword & " ")
    If pos = 0 Then EvalKeywordExpr = -1: Exit Function

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
    EvalKeywordExpr = EvalNumericExpr(exprStr)
End Function

Private Function ParseKeywordStringExpr(upperLine As String, originalLine As String, keyword As String) As String
    Dim pos As Integer
    pos = InStr(upperLine, " " & keyword & " ")
    If pos = 0 Then ParseKeywordStringExpr = "": Exit Function

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
    ParseKeywordStringExpr = EvalStringExpr(exprStr)
End Function


Private Sub SyncSelectionToPowerPoint(shapes As Collection, lineNum As Integer)
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
    Log "Line " & lineNum & ": WARNING - Could not sync selection to PowerPoint: " & Err.Description
End Sub


Private Function ExtractQuotedValue(s As String) As String
    Dim startQ As Integer
    startQ = InStr(s, """")
    If startQ = 0 Then
        Dim eqPos As Integer
        eqPos = InStr(s, "=")
        If eqPos > 0 Then
            ExtractQuotedValue = Trim(Mid(s, eqPos + 1))
        Else
            ExtractQuotedValue = Trim(s)
        End If
        Exit Function
    End If
    Dim endQ As Integer
    endQ = InStr(startQ + 1, s, """")
    If endQ = 0 Then
        ExtractQuotedValue = Mid(s, startQ + 1)
    Else
        ExtractQuotedValue = Mid(s, startQ + 1, endQ - startQ - 1)
    End If
End Function

Private Function HexToRGB(hexColor As String) As Long
    Dim h As String
    h = Trim(Replace(hexColor, "#", ""))
    HexToRGB = RGB(CLng("&H" & left(h, 2)), CLng("&H" & Mid(h, 3, 2)), CLng("&H" & right(h, 2)))
End Function

Private Sub Log(msg As String)
    If ScriptLog Is Nothing Then Set ScriptLog = New Collection
    ScriptLog.Add msg
End Sub


