Attribute VB_Name = "ModuleTableDistributeWithGaps"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

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

Sub TableDistributeColumnsWithGaps()

    Set MyDocument = Application.ActiveWindow
    Dim totalWidth As Double
    Dim NumberOfColumnsToDistribute As Long
    totalWidth = 0
    NumberOfColumnsToDistribute = 0
     
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
        
        For ColsCount = 1 To .Columns.count
        
            For RowsCount = 1 To .rows.count
                
                If .cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                totalWidth = totalWidth + .Columns(ColsCount).width
                NumberOfColumnsToDistribute = NumberOfColumnsToDistribute + 1
                Exit For
                    
                End If
                    
                
                End If
                
            Next RowsCount
        Next ColsCount
        
        
        If NumberOfColumnsToDistribute > 0 Then
        For ColsCount = 1 To .Columns.count
        
            For RowsCount = 1 To .rows.count
                
                If .cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                .Columns(ColsCount).width = totalWidth / NumberOfColumnsToDistribute
                Exit For
                    
                End If
                    
                
                End If
                
            Next RowsCount
        Next ColsCount
        End If
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub

Sub TableDistributeRowsWithGaps()

    Set MyDocument = Application.ActiveWindow
    Dim totalHeight As Double
    Dim NumberOfRowsToDistribute As Long
    totalHeight = 0
    NumberOfRowsToDistribute = 0
     
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
        
        For RowsCount = 1 To .rows.count
           
            For ColsCount = 1 To .Columns.count
                
                If .cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                totalHeight = totalHeight + .rows(RowsCount).height
                NumberOfRowsToDistribute = NumberOfRowsToDistribute + 1
                Exit For
                    
                End If
                    
                
                End If
                
            Next ColsCount
        Next RowsCount
        
        
        If NumberOfRowsToDistribute > 0 Then
        
        For RowsCount = 1 To .rows.count
        
            For ColsCount = 1 To .Columns.count
                
                If .cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                .rows(RowsCount).height = totalHeight / NumberOfRowsToDistribute
                Exit For
                    
                End If
                    
                
                End If
                
            Next ColsCount
        Next RowsCount
        End If
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub
