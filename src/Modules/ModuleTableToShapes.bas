Attribute VB_Name = "ModuleTableToShapes"
Sub ConvertTableToShapes()
    
    Set myDocument = Application.ActiveWindow
    
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
                .TextFrame.TextRange.Text = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text
                .TextFrame.TextRange.ParagraphFormat.Alignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.Alignment
                .TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment
                .TextFrame.TextRange.Font.Name = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Font.Name
                .TextFrame.TextRange.Font.Size = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Font.Size
                .TextFrame.TextRange.Font.Color = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Font.Color
                .TextFrame.TextRange.ParagraphFormat.Bullet = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.Bullet
                .Fill.ForeColor = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.Fill.ForeColor
                .Line.ForeColor = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders(ppBorderBottom).ForeColor
            End With
            
            TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            
        Next ColsCount
        
        TableLeft = Application.ActiveWindow.Selection.ShapeRange.Left
        TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
        
    Next RowsCount
    
    Application.ActiveWindow.Selection.ShapeRange.Delete
    
    
End Sub
