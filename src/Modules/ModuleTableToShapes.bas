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

Sub ConvertShapesToTable()
    Dim ConvertTable         As Table
    Dim numRows     As Integer
    Dim numCols     As Integer
    Dim minLeft     As Double
    Dim minTop      As Double
    Dim maxRight    As Double
    Dim maxBottom   As Double
    Dim cellShapes()    As shape
    Dim i As Long, j As Long
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    If MyDocument.Selection.ShapeRange.Count > 0 Then
        
        
        ObjectsQuicksortTopLeftToBottomRight MyDocument.Selection.ShapeRange
        
        ReDim cellShapes(1 To MyDocument.Selection.ShapeRange.Count)
        For i = 1 To MyDocument.Selection.ShapeRange.Count
            Set cellShapes(i) = MyDocument.Selection.ShapeRange(i)
        Next i
        minLeft = 99999
        minTop = 99999
        maxRight = 0
        maxBottom = 0
        
        
        'For i = 1 To UBound(cellShapes) - 1
        '    For j = i + 1 To UBound(cellShapes)
        '        If cellShapes(i).Top * 1000000 + cellShapes(i).left > cellShapes(j).Top * 1000000 + cellShapes(j).left Then
        '            Set temp = cellShapes(i)
        '            Set cellShapes(i) = cellShapes(j)
        '            Set cellShapes(j) = temp
        '        End If
        '    Next j
        'Next i
        
        For i = 1 To UBound(cellShapes)
            If cellShapes(i).left < minLeft Then minLeft = cellShapes(i).left
            If cellShapes(i).Top < minTop Then minTop = cellShapes(i).Top
            If cellShapes(i).left + cellShapes(i).Width > maxRight Then maxRight = cellShapes(i).left + cellShapes(i).Width
            If cellShapes(i).Top + cellShapes(i).Height > maxBottom Then maxBottom = cellShapes(i).Top + cellShapes(i).Height
        Next i
        
        'numCols based on user input, pre-calculated default setting
        numCols = Int(InputBox("Please specify the number of columns", "Number of columns", Int((maxRight - minLeft) / cellShapes(1).Width)))
        
        numRows = UBound(cellShapes) / numCols
        
        If (UBound(cellShapes) Mod numRows) > 0 Then
        numRows = numRows + 1
        End If
        
        For rowLoop = 1 To numRows
            
            If rowLoop * numCols > UBound(cellShapes) Then
                maxLoop = UBound(cellShapes)
            Else
                maxLoop = rowLoop * numCols
            End If
            
            For i = 1 + (rowLoop - 1) * (numCols) To maxLoop
                
                For j = i + 1 To maxLoop
                    
                    If (cellShapes(i).left > cellShapes(j).left) Then
                        Set temp = cellShapes(i)
                        Set cellShapes(i) = cellShapes(j)
                        Set cellShapes(j) = temp
                    End If
                Next j
            Next i
            
        Next rowLoop
        
        h = 1
        
        Set ConvertTable = ActiveWindow.Selection.SlideRange(1).Shapes.AddTable(numRows, numCols, minLeft, minTop).Table
        
        For i = 1 To UBound(cellShapes)
            
            j = ((i - 1) Mod numCols) + 1
            
               
                Set Newcell = ConvertTable.Cell(h, j)
                
                With Newcell.shape
                    .TextFrame.MarginBottom = cellShapes(i).TextFrame.MarginBottom
                    .TextFrame.MarginLeft = cellShapes(i).TextFrame.MarginLeft
                    .TextFrame.MarginRight = cellShapes(i).TextFrame.MarginRight
                    .TextFrame.MarginTop = cellShapes(i).TextFrame.MarginTop
                    
                    If cellShapes(i).TextFrame.HasText Then
                        cellShapes(i).TextFrame.TextRange.Copy
                        .TextFrame.TextRange.Paste
                    End If
                    
                    .TextFrame.TextRange.ParagraphFormat.Alignment = cellShapes(i).TextFrame.TextRange.ParagraphFormat.Alignment
                    .TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = cellShapes(i).TextFrame.TextRange.ParagraphFormat.BaseLineAlignment
                    .Fill.ForeColor.RGB = cellShapes(i).Fill.ForeColor.RGB
                    
                End With
                
                With Newcell
                    
                    If h = 1 Then
                        
                        If cellShapes(i).Line.Weight > -1 Then
                            .Borders(ppBorderBottom).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderTop).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderLeft).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderRight).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            
                            .Borders(ppBorderBottom).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderTop).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderLeft).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderRight).Weight = cellShapes(i).Line.Weight
                            
                            If cellShapes(i).Line.DashStyle > -1 Then
                                .Borders(ppBorderBottom).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderTop).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderLeft).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderRight).DashStyle = cellShapes(i).Line.DashStyle
                            End If
                            
                        Else
                            .Borders(ppBorderBottom).ForeColor.RGB = RGB(255, 255, 255)
                            
                            .Borders(ppBorderTop).ForeColor.RGB = RGB(255, 255, 255)
                            
                            If i = 1 Then
                                .Borders(ppBorderLeft).ForeColor.RGB = RGB(255, 255, 255)
                            End If
                            
                            .Borders(ppBorderRight).ForeColor.RGB = RGB(255, 255, 255)
                            
                            .Borders(ppBorderBottom).Weight = 0
                            .Borders(ppBorderTop).Weight = 0
                            If i = 1 Then
                                .Borders(ppBorderLeft).Weight = 0
                            End If
                            .Borders(ppBorderRight).Weight = 0
                        End If
                        
                    Else
                        
                        If cellShapes(i).Line.Weight > -1 Then
                            .Borders(ppBorderBottom).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderTop).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderLeft).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            .Borders(ppBorderRight).ForeColor.RGB = cellShapes(i).Line.ForeColor.RGB
                            
                            .Borders(ppBorderBottom).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderTop).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderLeft).Weight = cellShapes(i).Line.Weight
                            .Borders(ppBorderRight).Weight = cellShapes(i).Line.Weight
                            
                            If cellShapes(i).Line.DashStyle > -1 Then
                                .Borders(ppBorderBottom).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderTop).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderLeft).DashStyle = cellShapes(i).Line.DashStyle
                                .Borders(ppBorderRight).DashStyle = cellShapes(i).Line.DashStyle
                            End If
                            
                        Else
                            
                            .Borders(ppBorderBottom).Transparency = 0
                            .Borders(ppBorderRight).Transparency = 0
                            .Borders(ppBorderBottom).Weight = 0
                            .Borders(ppBorderRight).Weight = 0
                            
                        End If
                        
                    End If
                    
                End With
                
           
            If (j + 1) > numCols And Not i = UBound(cellShapes) Then
                h = h + 1
            End If
        
        cellShapes(i).Delete
        
        Next i
    Else
        MsgBox "No shapes selected."
    End If
    
    End If
End Sub



Sub ConvertTableToShapes()
    
    Set MyDocument = Application.ActiveWindow
            
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
    MsgBox "Please select a table."
    
    ElseIf MyDocument.Selection.ShapeRange.HasTable Then
    
    TableTop = MyDocument.Selection.ShapeRange.Top
    TableLeft = MyDocument.Selection.ShapeRange.left
    
    TypeOfColumnGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
    TypeOfRowGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
    
    ProgressForm.Show
    
    For RowsCount = 1 To MyDocument.Selection.ShapeRange.Table.Rows.Count
    
    SetProgress (RowsCount / MyDocument.Selection.ShapeRange.Table.Rows.Count * 100)
    
        For ColsCount = 1 To MyDocument.Selection.ShapeRange.Table.Columns.Count
            
            If Not ((ColsCount Mod 2 = 0 And TypeOfColumnGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfColumnGaps = "odd") Or (RowsCount Mod 2 = 0 And TypeOfRowGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfRowGaps = "odd")) Then
            
            Set NewShape = MyDocument.Selection.SlideRange.Shapes.AddShape(Type:=msoShapeRectangle, left:=TableLeft, Top:=TableTop, Width:=MyDocument.Selection.ShapeRange.Table.Columns(ColsCount).Width, Height:=MyDocument.Selection.ShapeRange.Table.Rows(RowsCount).Height)
            
            With NewShape
                .TextFrame.MarginBottom = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.MarginBottom
                .TextFrame.MarginLeft = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.MarginLeft
                .TextFrame.MarginRight = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.MarginRight
                .TextFrame.MarginTop = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.MarginTop
                
                If Not MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.Text = "" Then
                    MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.Cut
                    .TextFrame.TextRange.Paste
                End If
                
                .TextFrame.TextRange.ParagraphFormat.Alignment = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.ParagraphFormat.Alignment
                .TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment
                .Fill.ForeColor.RGB = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).shape.Fill.ForeColor.RGB
                .Line.ForeColor.RGB = MyDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders(ppBorderBottom).ForeColor.RGB
            End With
            
            End If
            
            TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            
        Next ColsCount
        
        
        TableLeft = Application.ActiveWindow.Selection.ShapeRange.left
        TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
        
    Next RowsCount
    
    ProgressForm.Hide
    Unload ProgressForm
    
    Application.ActiveWindow.Selection.ShapeRange.Delete
    
    Else
    
    MsgBox "No table selected."
    
    End If
       
End Sub
