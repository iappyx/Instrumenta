Attribute VB_Name = "ModuleTableSum"
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

Sub TableSum()

    Set MyDocument = Application.ActiveWindow
    Dim TotalSum As Double
    TotalSum = 0
    
       
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
    MsgBox "No table or cells selected."
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                    For SumCount = 1 To RowsCount - 1
                    
                    On Error Resume Next
                    TotalSum = TotalSum + CDbl(.Cell(SumCount, ColsCount).shape.TextFrame.TextRange.Text)
                    On Error GoTo 0
                    
                    Next SumCount
                        
                    .Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.Text = TotalSum
                    
                    End If
                    
                    TotalSum = 0
                
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub

Sub TableRowSum()

    Set MyDocument = Application.ActiveWindow
    Dim TotalSum As Double
    TotalSum = 0
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
    MsgBox "No table or cells selected."
    Else
    
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                    For SumCount = 1 To ColsCount - 1
                    
                    On Error Resume Next
                    TotalSum = TotalSum + CDbl(.Cell(RowsCount, SumCount).shape.TextFrame.TextRange.Text)
                    On Error GoTo 0
                    
                    Next SumCount
                        
                    .Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.Text = TotalSum
                    
                    End If
                    
                    TotalSum = 0
                
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub
