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
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).shape.TextFrame
                        
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
    
    End If
    
End Sub

Sub TablesMarginsIncrease()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
    Dim TableMarginSetting As Double
    TableMarginSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeMargin", "0" + GetDecimalSeperator() + "2"))
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).shape.TextFrame
                        
                        .MarginBottom = .MarginBottom + TableMarginSetting
                        .MarginLeft = .MarginLeft + TableMarginSetting
                        .MarginRight = .MarginRight + TableMarginSetting
                        .MarginTop = .MarginTop + TableMarginSetting
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If
    
End Sub

Sub TablesMarginsDecrease()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
    Dim TableMarginSetting As Double
    TableMarginSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeMargin", "0" + GetDecimalSeperator() + "2"))
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).shape.TextFrame
                        
                        If .MarginBottom >= TableMarginSetting Then
                            .MarginBottom = .MarginBottom - TableMarginSetting
                        End If
                        If .MarginLeft >= TableMarginSetting Then
                            .MarginLeft = .MarginLeft - TableMarginSetting
                        End If
                        If .MarginRight >= TableMarginSetting Then
                            .MarginRight = .MarginRight - TableMarginSetting
                        End If
                        If .MarginTop >= TableMarginSetting Then
                            .MarginTop = .MarginTop - TableMarginSetting
                        End If
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If
End Sub
