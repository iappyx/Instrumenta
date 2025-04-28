Attribute VB_Name = "ModuleTableInsertColumns"
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

Sub InsertColumnToLeftKeepOtherColumnWidths()
InsertColumnKeepOtherColumnWidths ("Left")
End Sub

Sub InsertColumnToRightKeepOtherColumnWidths()
InsertColumnKeepOtherColumnWidths ("Right")
End Sub

Sub InsertColumnKeepOtherColumnWidths(Position As String)
    
    Set MyDocument = Application.ActiveWindow
    If MyDocument.Selection.Type <> ppSelectionShapes And MyDocument.Selection.Type <> ppSelectionText Then
        MsgBox "No table or cells selected."
        Exit Sub
    End If
    
    If MyDocument.Selection.ShapeRange(1).HasTable Then
        
        For selectedRow = 1 To MyDocument.Selection.ShapeRange(1).table.Rows.Count
            For selectedColumn = 1 To MyDocument.Selection.ShapeRange(1).table.Columns.Count
                If MyDocument.Selection.ShapeRange(1).table.Cell(selectedRow, selectedColumn).Selected Then
                    Exit For
                End If
            Next selectedColumn
            If selectedColumn <= MyDocument.Selection.ShapeRange(1).table.Columns.Count Then Exit For
        Next selectedRow
        
        Dim newColumnWidth As Double
        newColumnWidth = MyDocument.Selection.ShapeRange(1).table.Columns(selectedColumn).Width
        
        Dim originalNumColumns As Integer
        originalNumColumns = MyDocument.Selection.ShapeRange(1).table.Columns.Count
        ReDim originalWidths(1 To originalNumColumns + 1)
        
        Dim offset As Integer
        For i = 1 To originalNumColumns
            If Position = "Left" And i >= selectedColumn Then
                offset = 1
            ElseIf Position = "Right" And i >= selectedColumn + 1 Then
                offset = 1
            Else
                offset = 0
            End If
            originalWidths(i + offset) = MyDocument.Selection.ShapeRange(1).table.Columns(i).Width
        Next i
        
        If Position = "Left" Then
            originalWidths(selectedColumn) = newColumnWidth
            MyDocument.Selection.ShapeRange(1).table.Columns.Add BeforeColumn:=selectedColumn
        ElseIf Position = "Right" Then
            originalWidths(selectedColumn + 1) = newColumnWidth
            MyDocument.Selection.ShapeRange(1).table.Columns.Add BeforeColumn:=selectedColumn + 1
        Else
            MsgBox "Invalid Position! Use 'Left' or 'Right'."
            Exit Sub
        End If
        
        For i = 1 To originalNumColumns + 1
            MyDocument.Selection.ShapeRange(1).table.Columns(i).Width = originalWidths(i)
        Next i
    Else
        MsgBox "No table or cells selected."
    End If
End Sub

