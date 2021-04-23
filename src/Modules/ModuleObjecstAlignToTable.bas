Attribute VB_Name = "ModuleObjecstAlignToTable"
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

Sub ObjectsAlignToTableColumn()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange(1).HasTable = True Then
        
        TableTop = myDocument.Selection.ShapeRange(1).Top
        TableLeft = myDocument.Selection.ShapeRange(1).Left
        
        AlignmentColumn = CInt(InputBox("Enter column number to align shapes to:", "Align objects to column", 1))
        SkipRows = CInt(InputBox("Enter first number of rows to skip (e.g. 1 if your table has a header row):", "Align objects to column", 1))
        
        For RowsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Rows.Count
            
            For ColsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Columns.Count
                
                If (ColsCount = AlignmentColumn) And RowsCount > SkipRows And (RowsCount < (myDocument.Selection.ShapeRange.Count + SkipRows)) Then
                    
                    With myDocument.Selection.ShapeRange(RowsCount + 1 - SkipRows)
                        
                        .Left = TableLeft + myDocument.Selection.ShapeRange(1).Table.Columns(ColsCount).Width / 2 - .Width / 2
                        .Top = TableTop + myDocument.Selection.ShapeRange(1).Table.Rows(RowsCount).Height / 2 - .Height / 2
                        
                    End With
                    
                End If
                
                TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange(1).Table.Columns(ColsCount).Width
                
            Next ColsCount
            
            TableLeft = Application.ActiveWindow.Selection.ShapeRange(1).Left
            TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange(1).Table.Rows(RowsCount).Height
            
        Next RowsCount
        
    Else
        
        MsgBox "First selected object is not a table. First select the table and then the shapes you want to align to a column of that table."
        
    End If
    
End Sub

Sub ObjectsAlignToTableRow()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange(1).HasTable = True Then
        
        TableTop = myDocument.Selection.ShapeRange(1).Top
        TableLeft = myDocument.Selection.ShapeRange(1).Left
        
        AlignmentRow = CInt(InputBox("Enter row number to align shapes to:", "Align objects to row", 1))
        SkipColumns = CInt(InputBox("Enter first number of columns to skip:", "Align objects to row", 0))
        
        For RowsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Rows.Count
            
            For ColsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Columns.Count
                
                If (RowsCount = AlignmentRow) And ColsCount > SkipColumns And (ColsCount < (myDocument.Selection.ShapeRange.Count + SkipColumns)) Then
                    
                    With myDocument.Selection.ShapeRange(ColsCount + 1 - SkipColumns)
                        
                        .Left = TableLeft + myDocument.Selection.ShapeRange(1).Table.Columns(ColsCount).Width / 2 - .Width / 2
                        .Top = TableTop + myDocument.Selection.ShapeRange(1).Table.Rows(RowsCount).Height / 2 - .Height / 2
                        
                    End With
                    
                End If
                
                TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange(1).Table.Columns(ColsCount).Width
                
            Next ColsCount
            
            TableLeft = Application.ActiveWindow.Selection.ShapeRange(1).Left
            TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange(1).Table.Rows(RowsCount).Height
            
        Next RowsCount
        
    Else
        
        MsgBox "First selected object is not a table. First select the table and then the shapes you want to align to a row of that table."
        
    End If
    
End Sub
