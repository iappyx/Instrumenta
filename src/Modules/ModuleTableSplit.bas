Attribute VB_Name = "ModuleTableSplit"
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

Sub SplitTableByRow()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
        
        If Application.ActiveWindow.Selection.ShapeRange(1).HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange(1).table
                
                For RowsCount = 1 To .rows.count
                    For ColsCount = 1 To .Columns.count
                        
                        If .cell(RowsCount, ColsCount).Selected Then
                            
                            If Not RowsCount = 1 Then
                                
                                Set ThisTable = Application.ActiveWindow.Selection.ShapeRange(1)
                                Set DuplicatedTable = ThisTable.Duplicate
                                DuplicatedTable.left = ThisTable.left
                                DuplicatedTable.Top = ThisTable.Top
                                
                                DuplicatedTable.table.FirstRow = False
                                
                                For DeleteRows = 1 To RowsCount - 1
                                    DuplicatedHeight = DuplicatedTable.table.rows(1).height
                                    DuplicatedTable.table.rows(1).Delete
                                    DuplicatedTable.Top = DuplicatedTable.Top + DuplicatedHeight
                                    
                                Next
                                
                                DuplicatedTable.Top = DuplicatedTable.Top + 5
                                
                                For DeleteRows = .rows.count To RowsCount Step -1
                                    ThisTable.table.rows(DeleteRows).Delete
                                Next
                                
                                Exit Sub
                                
                            Else
                                
                                MsgBox "Will not work on the first row."
                                
                            End If
                            
                        End If
                        
                    Next ColsCount
                Next RowsCount
                
            End With
            
        Else
            
            MsgBox "No table or cells selected."
            
        End If
        
    End If
    
End Sub

Sub SplitTableByColumn()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
        
        If Application.ActiveWindow.Selection.ShapeRange(1).HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange(1).table
                
                For RowsCount = 1 To .rows.count
                    For ColsCount = 1 To .Columns.count
                        
                        If .cell(RowsCount, ColsCount).Selected Then
                            
                            If Not ColsCount = 1 Then
                                
                                Set ThisTable = Application.ActiveWindow.Selection.ShapeRange(1)
                                Set DuplicatedTable = ThisTable.Duplicate
                                DuplicatedTable.left = ThisTable.left
                                DuplicatedTable.Top = ThisTable.Top
                                
                                DuplicatedTable.table.FirstCol = False
                                
                                For DeleteColumns = 1 To ColsCount - 1
                                    DuplicatedWidth = DuplicatedTable.table.Columns(1).width
                                    DuplicatedTable.table.Columns(1).Delete
                                    DuplicatedTable.left = DuplicatedTable.left + DuplicatedWidth
                                    
                                Next
                                
                                DuplicatedTable.left = DuplicatedTable.left + 5
                                
                                For DeleteColumns = .Columns.count To ColsCount Step -1
                                    ThisTable.table.Columns(DeleteColumns).Delete
                                Next
                                
                                Exit Sub
                                
                            Else
                                
                                MsgBox "Will not work on the first column."
                                
                            End If
                            
                        End If
                        
                    Next ColsCount
                Next RowsCount
                
            End With
            
        Else
            
            MsgBox "No table or cells selected."
            
        End If
        
    End If
    
End Sub
