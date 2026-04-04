Attribute VB_Name = "ModuleObjecstAlignToTable"
'MIT License

'Copyright (c) 2021 - 2026 iappyx
'
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

Sub ObjectsAlignToTable()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count > 1 Then
        
        Dim ShapeCount, TableIndex, TableDimensions  As Long
        
        TableIndex = 0
        TableDimensions = 0
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
            If MyDocument.Selection.ShapeRange(ShapeCount).HasTable = True Then
                
                If (MyDocument.Selection.ShapeRange(ShapeCount).width * MyDocument.Selection.ShapeRange(ShapeCount).height) > TableDimensions Then
                    TableIndex = ShapeCount
                    TableDimensions = MyDocument.Selection.ShapeRange(ShapeCount).width * MyDocument.Selection.ShapeRange(ShapeCount).height
                    
                End If
                
            End If
        Next ShapeCount
        
        If TableIndex >= 1 Then
            
            Dim SlideShape() As shape
            ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.count - 1)
            Dim SlideShapeCounter As Integer
            
            SlideShapeCounter = 1
            
            For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
                If ShapeCount <> TableIndex Then
                    Set SlideShape(SlideShapeCounter) = MyDocument.Selection.ShapeRange(ShapeCount)
                    SlideShapeCounter = SlideShapeCounter + 1
                End If
            Next ShapeCount
            
            Dim rows    As Integer, Columns As Integer
            
            rows = MyDocument.Selection.ShapeRange(TableIndex).table.rows.count
            Columns = MyDocument.Selection.ShapeRange(TableIndex).table.Columns.count
            
            Dim TableXBorder, TableYBorder, TableCols(), TableRows(), TableXCenter(), TableYCenter() As Double
            
            ReDim TableCols(Columns), TableRows(rows)
            ReDim TableXCenter(1 To Columns), TableYCenter(1 To rows)
            
            TableXBorder = MyDocument.Selection.ShapeRange(TableIndex).left
            TableYBorder = MyDocument.Selection.ShapeRange(TableIndex).Top
            
            TableRows(0) = TableYBorder
            TableCols(0) = TableXBorder
            
            For RowsCount = 1 To rows
                TableYBorder = TableYBorder + MyDocument.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height
                TableRows(RowsCount) = TableYBorder
                TableYCenter(RowsCount) = TableYBorder - (MyDocument.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height / 2)
            Next RowsCount
            
            For ColsCount = 1 To Columns
                TableXBorder = TableXBorder + MyDocument.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width
                TableCols(ColsCount) = TableXBorder
                TableXCenter(ColsCount) = TableXBorder - (MyDocument.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width / 2)
            Next ColsCount
            
            For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count - 1
                
                ShapeXCenter = SlideShape(ShapeCount).left + (SlideShape(ShapeCount).width / 2)
                ShapeYCenter = SlideShape(ShapeCount).Top + (SlideShape(ShapeCount).height / 2)
                
                For RowsCount = 1 To rows
                    
                    If ShapeYCenter >= TableRows(RowsCount - 1) And ShapeYCenter < TableRows(RowsCount) Then
                        
                        SlideShape(ShapeCount).Top = TableYCenter(RowsCount) - (SlideShape(ShapeCount).height / 2)
                        Exit For
                        
                    End If
                Next RowsCount
                
                For ColsCount = 1 To Columns
                    If ShapeXCenter >= TableCols(ColsCount - 1) And ShapeXCenter < TableCols(ColsCount) Then
                        
                        SlideShape(ShapeCount).left = TableXCenter(ColsCount) - (SlideShape(ShapeCount).width / 2)
                        Exit For
                        
                    End If
                Next ColsCount
            Next ShapeCount
            
        Else
            
            MsgBox "No table selected. Please Select a table."
            
        End If
        
    Else
    
        MsgBox "Select a table and some shapes."
        
    End If
    
End Sub

Sub ObjectsAlignToTableColumn()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count > 1 Then
        
        Dim ShapeCount, TableIndex, TableDimensions  As Long
        
        TableIndex = 0
        TableDimensions = 0
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
            If MyDocument.Selection.ShapeRange(ShapeCount).HasTable = True Then
                
                If MyDocument.Selection.ShapeRange(ShapeCount).height > TableDimensions Then
                    TableIndex = ShapeCount
                    TableDimensions = MyDocument.Selection.ShapeRange(ShapeCount).height
                    
                End If
                
            End If
        Next ShapeCount
        
        If TableIndex >= 1 Then
            
            TableTop = MyDocument.Selection.ShapeRange(TableIndex).Top
            TableLeft = MyDocument.Selection.ShapeRange(TableIndex).left
            
            AlignmentColumn = CInt(InputBox("Enter column number To align shapes to:", "Align objects To column", 1))
            SkipRows = CInt(InputBox("Enter first number of rows To skip (e.g. 1 If your table has a header row):", "Align objects To column", 1))
            
            SortOrder = MsgBox("Do you want use the top position of the shapes?" & vbNewLine & vbNewLine & "If you click Yes the top position of the shape will be used To distribute the different shapes" & vbNewLine & vbNewLine & "If you click No the order of selecting the different shapes will be used To distribute the different shapes", vbYesNo + vbQuestion, "Align objects To column")
            
            Dim SlideShape() As shape
            Dim SlideShapeOrdered() As shape
            ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.count - 1)
            ReDim SlideShapeOrdered(1 To MyDocument.Selection.ShapeRange.count)
            
            shapes = 0
            
            For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
                
                If ShapeCount = TableIndex Then
                    
                Else
                    shapes = shapes + 1
                    Set SlideShape(shapes) = MyDocument.Selection.ShapeRange(ShapeCount)
                    
                End If
                
            Next ShapeCount
            
            If SortOrder = vbYes Then
                ObjectsSortByTopPosition SlideShape
            End If
            
            Set SlideShapeOrdered(1) = MyDocument.Selection.ShapeRange(TableIndex)
            For ShapeCount = 2 To MyDocument.Selection.ShapeRange.count
                Set SlideShapeOrdered(ShapeCount) = SlideShape(ShapeCount - 1)
            Next ShapeCount
            
            For RowsCount = 1 To MyDocument.Selection.ShapeRange(TableIndex).table.rows.count
                
                For ColsCount = 1 To MyDocument.Selection.ShapeRange(TableIndex).table.Columns.count
                    
                    If (ColsCount = AlignmentColumn) And RowsCount > SkipRows And (RowsCount < (MyDocument.Selection.ShapeRange.count + SkipRows)) Then
                        
                        With SlideShapeOrdered(RowsCount + 1 - SkipRows)
                            
                            .left = TableLeft + MyDocument.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width / 2 - .width / 2
                            .Top = TableTop + MyDocument.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height / 2 - .height / 2
                            
                        End With
                        
                    End If
                    
                    TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width
                    
                Next ColsCount
                
                TableLeft = Application.ActiveWindow.Selection.ShapeRange(TableIndex).left
                TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height
                
            Next RowsCount
            
        Else
            
            MsgBox "The selection contains no table."
            
        End If
        
    Else
    
        MsgBox "Select a table and some shapes."
        
    End If
    
End Sub
Sub ObjectsAlignToTableRow()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count > 1 Then
        
        Dim ShapeCount, TableIndex, TableDimensions  As Long
        
        TableIndex = 0
        TableDimensions = 0
        
        For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
            If MyDocument.Selection.ShapeRange(ShapeCount).HasTable = True Then
                
                If MyDocument.Selection.ShapeRange(ShapeCount).width > TableDimensions Then
                    TableIndex = ShapeCount
                    TableDimensions = MyDocument.Selection.ShapeRange(ShapeCount).width
                    
                End If
                
            End If
        Next ShapeCount
        
        If TableIndex >= 1 Then
            
            TableTop = MyDocument.Selection.ShapeRange(TableIndex).Top
            TableLeft = MyDocument.Selection.ShapeRange(TableIndex).left
            
            AlignmentRow = CInt(InputBox("Enter row number To align shapes to:", "Align objects To row", 1))
            SkipColumns = CInt(InputBox("Enter first number of columns To skip:", "Align objects To row", 0))
            
            SortOrder = MsgBox("Do you want use the left position of the shapes?" & vbNewLine & vbNewLine & "If you click Yes the left position of the shape will be used To distribute the different shapes" & vbNewLine & vbNewLine & "If you click No the order of selecting the different shapes will be used To distribute the different shapes", vbYesNo + vbQuestion, "Align objects To row")
            
            Dim SlideShape() As shape
            Dim SlideShapeOrdered() As shape
            ReDim SlideShape(1 To MyDocument.Selection.ShapeRange.count - 1)
            ReDim SlideShapeOrdered(1 To MyDocument.Selection.ShapeRange.count)
            
            shapes = 0
            
            For ShapeCount = 1 To MyDocument.Selection.ShapeRange.count
                
                If ShapeCount = TableIndex Then
                    
                Else
                    shapes = shapes + 1
                    Set SlideShape(shapes) = MyDocument.Selection.ShapeRange(ShapeCount)
                    
                End If
                
            Next ShapeCount
            
            If SortOrder = vbYes Then
                ObjectsSortByLeftPosition SlideShape
            End If
            
            Set SlideShapeOrdered(1) = MyDocument.Selection.ShapeRange(TableIndex)
            For ShapeCount = 2 To MyDocument.Selection.ShapeRange.count
                Set SlideShapeOrdered(ShapeCount) = SlideShape(ShapeCount - 1)
            Next ShapeCount
            
            For RowsCount = 1 To MyDocument.Selection.ShapeRange(TableIndex).table.rows.count
                
                For ColsCount = 1 To MyDocument.Selection.ShapeRange(TableIndex).table.Columns.count
                    
                    If (RowsCount = AlignmentRow) And ColsCount > SkipColumns And (ColsCount < (MyDocument.Selection.ShapeRange.count + SkipColumns)) Then
                        
                        With SlideShapeOrdered(ColsCount + 1 - SkipColumns)
                            
                            .left = TableLeft + MyDocument.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width / 2 - .width / 2
                            .Top = TableTop + MyDocument.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height / 2 - .height / 2
                            
                        End With
                        
                    End If
                    
                    TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange(TableIndex).table.Columns(ColsCount).width
                    
                Next ColsCount
                
                TableLeft = Application.ActiveWindow.Selection.ShapeRange(TableIndex).left
                TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange(TableIndex).table.rows(RowsCount).height
                
            Next RowsCount
            
        Else
            
            MsgBox "The selection contains no table."
            
        End If
        
    Else
    
        MsgBox "Select a table and some shapes."
        
    End If
    
End Sub


