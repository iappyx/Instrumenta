Attribute VB_Name = "ModuleInstrumentaScript_Expr"
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

Public Function IScr_ComputeNumber(expr As String) As Double

    Dim substituted As String
    substituted = IScr_SubstituteVars(expr)

    IScr_exprTokens = IScr_SplitExpression(substituted)
    IScr_exprPos = 0

    On Error GoTo Failed
    IScr_ComputeNumber = IScr_ReadExpressionAddSub()
    Exit Function
Failed:
    IScr_ComputeNumber = 0
End Function

Public Function IScr_SubstituteVars(expr As String) As String
    Dim result As String
    result = expr
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        If IScr_varIsString(i) Then
            
        Else
            result = IScr_ReplaceWholeWord(result, IScr_varNames(i), CStr(IScr_varValues(i)))
        End If
    Next i
    IScr_SubstituteVars = result
End Function

Public Function IScr_SplitExpression(expr As String) As String()
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

    IScr_SplitExpression = tokens
End Function

Public Function IScr_ReadExpressionAddSub() As Double
    Dim left As Double
    left = IScr_ReadExpressionMulDiv()

    Do While IScr_exprPos <= UBound(IScr_exprTokens)
        Dim op As String
        op = IScr_exprTokens(IScr_exprPos)
        If op = "+" Or op = "-" Then
            IScr_exprPos = IScr_exprPos + 1
            Dim right As Double
            right = IScr_ReadExpressionMulDiv()
            If op = "+" Then left = left + right Else left = left - right
        Else
            Exit Do
        End If
    Loop

    IScr_ReadExpressionAddSub = left
End Function

Public Function IScr_ReadExpressionMulDiv() As Double
    Dim left As Double
    left = IScr_ReadExpressionUnary()

    Do While IScr_exprPos <= UBound(IScr_exprTokens)
        Dim op As String
        op = IScr_exprTokens(IScr_exprPos)
        If op = "*" Or op = "/" Then
            IScr_exprPos = IScr_exprPos + 1
            Dim right As Double
            right = IScr_ReadExpressionUnary()
            If op = "*" Then
                left = left * right
            Else
                If right <> 0 Then left = left / right Else left = 0
            End If
        Else
            Exit Do
        End If
    Loop

    IScr_ReadExpressionMulDiv = left
End Function

Public Function IScr_ReadExpressionUnary() As Double
    If IScr_exprPos <= UBound(IScr_exprTokens) Then
        If IScr_exprTokens(IScr_exprPos) = "-" Then
            IScr_exprPos = IScr_exprPos + 1
            IScr_ReadExpressionUnary = -IScr_ReadExpressionPrimary()
            Exit Function
        ElseIf IScr_exprTokens(IScr_exprPos) = "+" Then
            IScr_exprPos = IScr_exprPos + 1
        End If
    End If
    IScr_ReadExpressionUnary = IScr_ReadExpressionPrimary()
End Function

Public Function IScr_ReadExpressionPrimary() As Double
    If IScr_exprPos > UBound(IScr_exprTokens) Then
        IScr_ReadExpressionPrimary = 0
        Exit Function
    End If

    Dim tok As String
    tok = IScr_exprTokens(IScr_exprPos)

    If tok = "(" Then
        IScr_exprPos = IScr_exprPos + 1
        Dim val As Double
        val = IScr_ReadExpressionAddSub()
        If IScr_exprPos <= UBound(IScr_exprTokens) And IScr_exprTokens(IScr_exprPos) = ")" Then
            IScr_exprPos = IScr_exprPos + 1
        End If
        IScr_ReadExpressionPrimary = val
    Else
        IScr_exprPos = IScr_exprPos + 1
        IScr_ReadExpressionPrimary = CDbl(tok)
    End If
End Function


