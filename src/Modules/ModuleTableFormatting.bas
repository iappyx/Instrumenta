Attribute VB_Name = "ModuleTableFormatting"
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

Sub TableQuickFormat()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                TableRemoveBackgrounds
                TableRemoveBorders
                
                If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "" Then
                    TableColumnRemoveGaps
                End If
                
                If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "" Then
                    TableRowRemoveGaps
                End If
                
                For RowCount = 1 To .Rows.Count
                    
                    For ColumnCount = 1 To .Columns.Count
                        
                        .Cell(RowCount, ColumnCount).shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)

                    Next
                    
                Next
                
                For CellCount = 1 To .Rows(1).Cells.Count
                    
                    .Rows(1).Cells(CellCount).Borders(ppBorderTop).Weight = 0
                    .Rows(1).Cells(CellCount).Borders(ppBorderBottom).Weight = 2
                    .Rows(1).Cells(CellCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                    .Rows(1).Cells(CellCount).shape.Fill.visible = msoFalse
                    .Rows(1).Cells(CellCount).shape.TextFrame.VerticalAnchor = msoAnchorBottom
                    .Rows(1).Cells(CellCount).shape.TextFrame.TextRange.Font.Bold = msoTrue
                    .Rows(1).Cells(CellCount).shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                    
                Next CellCount
                
                TableColumnGaps "even", 20
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRemoveBackgrounds()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table

                .HorizBanding = False
                .VertBanding = False
                
                Application.ActiveWindow.Selection.ShapeRange.Fill.Solid
                Application.ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Application.ActiveWindow.Selection.ShapeRange.Fill.visible = msoFalse
                
                .Background.Fill.Solid
                .Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Background.Fill.visible = msoFalse
                
                ProgressForm.Show
                
                For RowCount = 1 To .Rows.Count
                    
                SetProgress (RowCount / .Rows.Count * 100)
                    
                    For ColumnCount = 1 To .Columns.Count
                        
                        .Cell(RowCount, ColumnCount).shape.Fill.Solid
                        .Cell(RowCount, ColumnCount).shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).shape.Fill.visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                Unload ProgressForm
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRemoveBorders()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
            
                ProgressForm.Show
                
                For RowCount = 1 To .Rows.Count
                    
                SetProgress (RowCount / .Rows.Count * 100)
                    
                    For ColumnCount = 1 To .Columns.Count
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(255, 255, 255)
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).Weight = 0
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                Unload ProgressForm
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even" Then
                
                If MsgBox("Existing column gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                    TableColumnRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                NumberOfColumns = .Columns.Count
                Dim ColumnWidthArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns + 1
                    ReDim ColumnWidthArray(0)
                    
                    For ColumnCount = 1 To NumberOfColumns
                        ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                        ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                        
                        If ColumnCount = NumberOfColumns Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = GapSize
                        End If
                        
                    Next ColumnCount
                    
                Else
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns - 1
                    
                    For ColumnCount = 1 To NumberOfColumns
                        
                        If Not ColumnCount = 1 Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                            ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                            
                        Else
                            ReDim ColumnWidthArray(1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        End If
                        
                    Next ColumnCount
                    
                End If
                
                For ColumnCount = NumberOfColumns To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedColumn = .Columns.Add(ColumnCount)
                        
                        For CellCount = 1 To AddedColumn.Cells.Count
                            AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                            AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                            AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                            
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If ColumnCount = NumberOfColumns Then
                            
                            Set AddedColumn = .Columns.Add
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not ColumnCount = 1 Then
                            
                            Set AddedColumn = .Columns.Add(ColumnCount)
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next ColumnCount
                
                For ColumnCount = 1 To NumberOfNewColumns
                    
                    .Columns(ColumnCount).Width = ColumnWidthArray(ColumnCount - 1)
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnIncreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Dim ColumnGapSetting As Double
            ColumnGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width + ColumnGapSetting
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnDecreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Dim ColumnGapSetting As Double
            ColumnGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If ((ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Columns(ColumnCount).Width - ColumnGapSetting) >= 0)) Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width - ColumnGapSetting
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnRemoveGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA COLUMNGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = .Columns.Count To 1 Step -1
                    
                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Delete
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even" Then
                
                If MsgBox("Existing row gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                    TableRowRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                NumberOfRows = .Rows.Count
                Dim RowHeightArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows + 1
                    ReDim RowHeightArray(0)
                    
                    For RowCount = 1 To NumberOfRows
                        ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                        RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                        RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                        
                        If RowCount = NumberOfRows Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = GapSize
                        End If
                        
                    Next RowCount
                    
                Else
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows - 1
                    
                    For RowCount = 1 To NumberOfRows
                        
                        If Not RowCount = 1 Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                            RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                            
                        Else
                            ReDim RowHeightArray(1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                        End If
                        
                    Next RowCount
                    
                End If
                
                For RowCount = NumberOfRows To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedRow = .Rows.Add(RowCount)
                        
                        For CellCount = 1 To AddedRow.Cells.Count
                            AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                            AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                            AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                            
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If RowCount = NumberOfRows Then
                            
                            Set AddedRow = .Rows.Add
                            
                            For CellCount = 1 To AddedRow.Cells.Count
                                AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not RowCount = 1 Then
                            
                            Set AddedRow = .Rows.Add(RowCount)
                            
                            For CellCount = 1 To AddedRow.Cells.Count
                                AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next RowCount
                
                For RowCount = 1 To NumberOfNewRows
                    
                    .Rows(RowCount).Height = RowHeightArray(RowCount - 1)
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowIncreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Dim RowGapSetting As Double
            RowGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height + RowGapSetting
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowDecreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Dim RowGapSetting As Double
            RowGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If ((RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Rows(RowCount).Height - RowGapSetting) >= 0)) Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height - RowGapSetting
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowRemoveGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA ROWGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = .Rows.Count To 1 Step -1
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Delete
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
