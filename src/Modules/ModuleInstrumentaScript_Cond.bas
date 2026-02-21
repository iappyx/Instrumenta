Attribute VB_Name = "ModuleInstrumentaScript_Cond"
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


Public Function IScr_EvalCondition(cond As String) As Boolean
    On Error GoTo Failed
    IScr_EvalCondition = IScr_ParseCondOr(Trim(cond))
    Exit Function
Failed:
    IScr_EvalCondition = False
End Function

Public Function IScr_ParseCondOr(cond As String) As Boolean
   
    Dim parts() As String
    parts = IScr_SplitCondOn(cond, " OR ")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If IScr_ParseCondAnd(Trim(parts(i))) Then
            IScr_ParseCondOr = True
            Exit Function
        End If
    Next i
    IScr_ParseCondOr = False
End Function

Public Function IScr_ParseCondAnd(cond As String) As Boolean
    
    Dim parts() As String
    parts = IScr_SplitCondOn(cond, " AND ")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If Not IScr_ParseCondNot(Trim(parts(i))) Then
            IScr_ParseCondAnd = False
            Exit Function
        End If
    Next i
    IScr_ParseCondAnd = True
End Function

Public Function IScr_ParseCondNot(cond As String) As Boolean
    Dim c As String
    c = Trim(cond)
    If left(UCase(c), 4) = "NOT " Then
        IScr_ParseCondNot = Not IScr_ParseCondAtom(Trim(Mid(c, 5)))
    Else
        IScr_ParseCondNot = IScr_ParseCondAtom(c)
    End If
End Function

Public Function IScr_ParseCondAtom(cond As String) As Boolean
    Dim c As String
    c = Trim(cond)

    If left(c, 1) = "(" And right(c, 1) = ")" Then
        IScr_ParseCondAtom = IScr_ParseCondOr(Trim(Mid(c, 2, Len(c) - 2)))
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
            leftVal = IScr_ComputeNumber(leftExpr)
            rightVal = IScr_ComputeNumber(rightExpr)

            Select Case opStr
                Case ">=": IScr_ParseCondAtom = (leftVal >= rightVal)
                Case "<=": IScr_ParseCondAtom = (leftVal <= rightVal)
                Case "<>": IScr_ParseCondAtom = (leftVal <> rightVal)
                Case ">": IScr_ParseCondAtom = (leftVal > rightVal)
                Case "<": IScr_ParseCondAtom = (leftVal < rightVal)
                Case "=": IScr_ParseCondAtom = (leftVal = rightVal)
            End Select
            Exit Function
        End If
    Next opIdx

    If UCase(c) = "TRUE" Then IScr_ParseCondAtom = True
    If UCase(c) = "FALSE" Then IScr_ParseCondAtom = False
End Function

Public Function IScr_SplitCondOn(cond As String, separator As String) As String()
   
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

    IScr_SplitCondOn = parts
End Function
