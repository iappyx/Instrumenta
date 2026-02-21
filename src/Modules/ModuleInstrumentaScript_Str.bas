Attribute VB_Name = "ModuleInstrumentaScript_Str"
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

Public Function IScr_SubstituteStringVars(expr As String) As String
    Dim result As String
    result = expr
    Dim i As Integer
    For i = 0 To IScr_varCount - 1
        
        Dim replacement As String
        If IScr_varIsString(i) Then
            replacement = """" & IScr_varStrValues(i) & """"
        Else
            If IScr_varValues(i) = Int(IScr_varValues(i)) Then
                replacement = CStr(CLng(IScr_varValues(i)))
            Else
                replacement = CStr(IScr_varValues(i))
            End If
        End If
        result = IScr_ReplaceWholeWordOutsideQuotes(result, IScr_varNames(i), replacement)
    Next i
    IScr_SubstituteStringVars = result
End Function

Public Function IScr_ReplaceWholeWord(s As String, word As String, replacement As String) As String
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
                    beforeOk = Not IScr_IsAlphaNum(Mid(s, i - 1, 1))
                End If

                If i + wLen > sLen Then
                    afterOk = True
                Else
                    afterOk = Not IScr_IsAlphaNum(Mid(s, i + wLen, 1))
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
    IScr_ReplaceWholeWord = result
End Function

Public Function IScr_ReplaceWholeWordOutsideQuotes(s As String, word As String, replacement As String) As String
    Dim result As String
    result = ""
    Dim i As Integer
    i = 1
    Dim inQuote As Boolean
    inQuote = False
    Dim sLen As Integer: sLen = Len(s)
    Dim wLen As Integer: wLen = Len(word)

    Do While i <= sLen
        Dim c As String
        c = Mid(s, i, 1)
        If c = """" Then
            inQuote = Not inQuote
            result = result & c
            i = i + 1
        ElseIf Not inQuote And i + wLen - 1 <= sLen Then
            If LCase(Mid(s, i, wLen)) = LCase(word) Then
                Dim beforeOk As Boolean
                Dim afterOk As Boolean
                If i = 1 Then
                    beforeOk = True
                Else
                    beforeOk = Not IScr_IsAlphaNum(Mid(s, i - 1, 1))
                End If
                If i + wLen > sLen Then
                    afterOk = True
                Else
                    afterOk = Not IScr_IsAlphaNum(Mid(s, i + wLen, 1))
                End If
                If beforeOk And afterOk Then
                    result = result & replacement
                    i = i + wLen
                    GoTo ContinueLoop
                End If
            End If
            result = result & c
            i = i + 1
        Else
            result = result & c
            i = i + 1
        End If
ContinueLoop:
    Loop
    IScr_ReplaceWholeWordOutsideQuotes = result
End Function

Public Function IScr_IsAlphaNum(c As String) As Boolean
    IScr_IsAlphaNum = (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or _
                 (c >= "0" And c <= "9") Or c = "_"
End Function



Public Function IScr_ComputeText(expr As String) As String
    Dim substituted As String
    substituted = IScr_SubstituteStringVars(expr)

    Dim parts() As String
    parts = IScr_SplitStringExpr(substituted)

    Dim result As String
    result = ""
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim part As String
        part = Trim(parts(i))
        If left(part, 1) = """" Then
            
            result = result & Mid(part, 2, Len(part) - 2)
        ElseIf IScr_IsNumericExpr(part) Then
            
            Dim numVal As Double
            numVal = IScr_ComputeNumber(part)
            
            If numVal = Int(numVal) Then
                result = result & CStr(CLng(numVal))
            Else
                result = result & CStr(numVal)
            End If
        Else
            result = result & part
        End If
    Next i

    IScr_ComputeText = result
End Function

Public Function IScr_SplitStringExpr(expr As String) As String()
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

    IScr_SplitStringExpr = parts
End Function

Public Function IScr_IsNumericExpr(s As String) As Boolean
    IScr_IsNumericExpr = (InStr(s, """") = 0)
End Function

