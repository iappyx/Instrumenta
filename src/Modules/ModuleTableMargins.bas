Attribute VB_Name = "ModuleTableMargins"
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


Sub TablesMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
End Sub

Sub TablesMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        .MarginBottom = .MarginBottom + 0.2
                        .MarginLeft = .MarginLeft + 0.2
                        .MarginRight = .MarginRight + 0.2
                        .MarginTop = .MarginTop + 0.2
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
End Sub

Sub TablesMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        If .MarginBottom >= 0.2 Then
                            .MarginBottom = .MarginBottom - 0.2
                        End If
                        If .MarginLeft >= 0.2 Then
                            .MarginLeft = .MarginLeft - 0.2
                        End If
                        If .MarginRight >= 0.2 Then
                            .MarginRight = .MarginRight - 0.2
                        End If
                        If .MarginTop >= 0.2 Then
                            .MarginTop = .MarginTop - 0.2
                        End If
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
End Sub
