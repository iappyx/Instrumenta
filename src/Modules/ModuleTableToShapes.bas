Attribute VB_Name = "ModuleTableToShapes"
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

Sub ConvertTableToShapes()
    
    Set myDocument = Application.ActiveWindow
            
    If Not myDocument.Selection.Type = ppSelectionShapes Then
    MsgBox "Please select a table."
    
    ElseIf myDocument.Selection.ShapeRange.HasTable Then
    
    TableTop = myDocument.Selection.ShapeRange.Top
    TableLeft = myDocument.Selection.ShapeRange.Left
    
    For RowsCount = 1 To myDocument.Selection.ShapeRange.Table.Rows.Count
        For ColsCount = 1 To myDocument.Selection.ShapeRange.Table.Columns.Count
            
            Set NewShape = myDocument.Selection.SlideRange.Shapes.AddShape(Type:=msoShapeRectangle, Left:=TableLeft, Top:=TableTop, Width:=myDocument.Selection.ShapeRange.Table.Columns(ColsCount).Width, Height:=myDocument.Selection.ShapeRange.Table.Rows(RowsCount).Height)
            
            With NewShape
                .TextFrame.MarginBottom = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginBottom
                .TextFrame.MarginLeft = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginLeft
                .TextFrame.MarginRight = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginRight
                .TextFrame.MarginTop = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginTop
                
                myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Cut
                .TextFrame.TextRange.Paste
                
                .TextFrame.TextRange.ParagraphFormat.Alignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.Alignment
                .TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment
                .Fill.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.Fill.ForeColor.RGB
                .Line.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders(ppBorderBottom).ForeColor.RGB
            End With
            
            TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            
        Next ColsCount
        
        TableLeft = Application.ActiveWindow.Selection.ShapeRange.Left
        TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
        
    Next RowsCount
    
    Application.ActiveWindow.Selection.ShapeRange.Delete
    
    Else
    
    MsgBox "No table selected."
    
    End If
       
End Sub
