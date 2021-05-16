Attribute VB_Name = "ModuleFunctions"
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

Function RemoveDuplicates(InputArray) As Variant
    
    Dim OutputArray, InputValue, OutputValue As Variant
    Dim MatchFound  As Boolean
    
    On Error Resume Next
    OutputArray = Array("")
    For Each InputValue In InputArray
        MatchFound = False
        
        If IsEmpty(InputValue) Then GoTo ForceNext
        For Each OutputValue In OutputArray
            If OutputValue = InputValue Then
                MatchFound = True
                Exit For
            End If
        Next OutputValue
        
        If MatchFound = False Then
            ReDim Preserve OutputArray(UBound(OutputArray, 1) + 1)
            OutputArray(UBound(OutputArray, 1) - 1) = InputValue
        End If
        
ForceNext:
    Next
    RemoveDuplicates = OutputArray
    
End Function
