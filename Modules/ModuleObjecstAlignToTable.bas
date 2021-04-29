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
        SkipRows = CInt(InputBox("Enter first number of rows to skip (e.g. 1 If your table has a header row):", "Align objects to column", 1))
        
        SortOrder = MsgBox("Do you want use the top position of the shapes?" & vbNewLine & vbNewLine & "If you click Yes the top position of the shape will be used to distribute the different shapes" & vbNewLine & vbNewLine & "If you click No the order of selecting the different shapes will be used to distribute the different shapes", vbYesNo + vbQuestion, "Align objects to column")
        
        Dim ShapeCount  As Long
        Dim SlideShape() As Shape
        Dim SlideShapeOrdered() As Shape
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count - 1)
        ReDim SlideShapeOrdered(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count - 1
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount + 1)
        Next ShapeCount
        
        If SortOrder = vbYes Then
            ObjectsSortByTopPosition SlideShape
        End If
        
        Set SlideShapeOrdered(1) = myDocument.Selection.ShapeRange(1)
        For ShapeCount = 2 To myDocument.Selection.ShapeRange.Count
            Set SlideShapeOrdered(ShapeCount) = SlideShape(ShapeCount - 1)
        Next ShapeCount
        
        For RowsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Rows.Count
            
            For ColsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Columns.Count
                
                If (ColsCount = AlignmentColumn) And RowsCount > SkipRows And (RowsCount < (myDocument.Selection.ShapeRange.Count + SkipRows)) Then
                    
                    With SlideShapeOrdered(RowsCount + 1 - SkipRows)
                        
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
        
        MsgBox "First selected object Is not a table. First select the table and then the shapes you want to align to a column of that table."
        
    End If
    
End Sub

Sub ObjectsAlignToTableRow()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange(1).HasTable = True Then
        
        TableTop = myDocument.Selection.ShapeRange(1).Top
        TableLeft = myDocument.Selection.ShapeRange(1).Left
        
        AlignmentRow = CInt(InputBox("Enter row number to align shapes to:", "Align objects to row", 1))
        SkipColumns = CInt(InputBox("Enter first number of columns to skip:", "Align objects to row", 0))
        
        SortOrder = MsgBox("Do you want use the left position of the shapes?" & vbNewLine & vbNewLine & "If you click Yes the left position of the shape will be used to distribute the different shapes" & vbNewLine & vbNewLine & "If you click No the order of selecting the different shapes will be used to distribute the different shapes", vbYesNo + vbQuestion, "Align objects to row")
        
        Dim ShapeCount  As Long
        Dim SlideShape() As Shape
        Dim SlideShapeOrdered() As Shape
        ReDim SlideShape(1 To myDocument.Selection.ShapeRange.Count - 1)
        ReDim SlideShapeOrdered(1 To myDocument.Selection.ShapeRange.Count)
        
        For ShapeCount = 1 To myDocument.Selection.ShapeRange.Count - 1
            Set SlideShape(ShapeCount) = myDocument.Selection.ShapeRange(ShapeCount + 1)
        Next ShapeCount
        
        If SortOrder = vbYes Then
            ObjectsSortByLeftPosition SlideShape
        End If
        
        Set SlideShapeOrdered(1) = myDocument.Selection.ShapeRange(1)
        For ShapeCount = 2 To myDocument.Selection.ShapeRange.Count
            Set SlideShapeOrdered(ShapeCount) = SlideShape(ShapeCount - 1)
        Next ShapeCount
        
        For RowsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Rows.Count
            
            For ColsCount = 1 To myDocument.Selection.ShapeRange(1).Table.Columns.Count
                
                If (RowsCount = AlignmentRow) And ColsCount > SkipColumns And (ColsCount < (myDocument.Selection.ShapeRange.Count + SkipColumns)) Then
                    
                    With SlideShapeOrdered(ColsCount + 1 - SkipColumns)
                        
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
        
        MsgBox "First selected Object Is Not a table. First Select the table And Then the shapes you want To align To a row of that table."
        
    End If
    
End Sub
