Attribute VB_Name = "ModuleTableMoveRowsOrColumns"
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

Sub MoveTableRowUpTextOnly()
Call MoveSelectedRowOrColumn("up", True, False)
End Sub

Sub MoveTableRowDownTextOnly()
Call MoveSelectedRowOrColumn("down", True, False)
End Sub

Sub MoveTableColumnRightTextOnly()
Call MoveSelectedRowOrColumn("right", True, False)
End Sub

Sub MoveTableColumnLeftTextOnly()
Call MoveSelectedRowOrColumn("left", True, False)
End Sub

Sub MoveTableRowUpIgnoreBorders()
Call MoveSelectedRowOrColumn("up", False, True)
End Sub

Sub MoveTableRowDownIgnoreBorders()
Call MoveSelectedRowOrColumn("down", False, True)
End Sub

Sub MoveTableColumnRightIgnoreBorders()
Call MoveSelectedRowOrColumn("right", False, True)
End Sub

Sub MoveTableColumnLeftIgnoreBorders()
Call MoveSelectedRowOrColumn("left", False, False)
End Sub

Sub MoveTableRowUp()
Call MoveSelectedRowOrColumn("up", False, False)
End Sub

Sub MoveTableRowDown()
Call MoveSelectedRowOrColumn("down", False, False)
End Sub

Sub MoveTableColumnRight()
Call MoveSelectedRowOrColumn("right", False, False)
End Sub

Sub MoveTableColumnLeft()
Call MoveSelectedRowOrColumn("left", False, False)
End Sub

Sub MoveSelectedRowOrColumn(moveDirection As String, textOnly As Boolean, ignoreBorders As Boolean)
    Dim MyDocument As Object
    Dim table As table
    Dim RowsCount As Integer, ColsCount As Integer
    Dim selectedRowIndex As Integer
    Dim selectedColIndex As Integer
    Dim i As Integer

    Set MyDocument = Application.ActiveWindow

    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
        Exit Sub
    End If

    If (MyDocument.Selection.ShapeRange.Count = 1) And MyDocument.Selection.ShapeRange.HasTable Then
        Set table = MyDocument.Selection.ShapeRange.table

        For RowsCount = 1 To table.Rows.Count
            For ColsCount = 1 To table.Columns.Count
                If table.Cell(RowsCount, ColsCount).Selected Then
                    selectedRowIndex = RowsCount
                    selectedColIndex = ColsCount
                    Exit For
                End If
            Next ColsCount
            If selectedRowIndex > 0 Then Exit For
        Next RowsCount

        If selectedRowIndex = 0 Or selectedColIndex = 0 Then
            MsgBox "Unable to detect selected cell. Ensure your cursor is inside a table cell.", vbCritical
            Exit Sub
        End If

        Set CopyTable = MyDocument.Selection.ShapeRange.Duplicate
        CopyTable.Width = Application.ActiveWindow.Selection.ShapeRange.Width
        CopyTable.Top = Application.ActiveWindow.Selection.ShapeRange.Top
        CopyTable.left = Application.ActiveWindow.Selection.ShapeRange.left

        With table
            Select Case moveDirection
                Case "up"
                    If selectedRowIndex > 1 Then
                        Call MoveRow(table, CopyTable, selectedRowIndex, selectedRowIndex - 1, textOnly, ignoreBorders)
                        MyDocument.Selection.ShapeRange.Delete
                        CopyTable.Select
                        CopyTable.table.Cell(selectedRowIndex - 1, selectedColIndex).Select
                    Else
                        MsgBox "Row is already at the top.", vbInformation
                    End If
                Case "down"
                    If selectedRowIndex < .Rows.Count Then
                        Call MoveRow(table, CopyTable, selectedRowIndex, selectedRowIndex + 1, textOnly, ignoreBorders)
                        MyDocument.Selection.ShapeRange.Delete
                        CopyTable.Select
                        CopyTable.table.Cell(selectedRowIndex + 1, selectedColIndex).Select
                    Else
                        MsgBox "Row is already at the bottom.", vbInformation
                    End If
                Case "left"
                    If selectedColIndex > 1 Then
                        Call MoveColumn(table, CopyTable, selectedColIndex, selectedColIndex - 1, textOnly, ignoreBorders)
                        MyDocument.Selection.ShapeRange.Delete
                        CopyTable.Select
                        CopyTable.table.Cell(selectedRowIndex, selectedColIndex - 1).Select
                    Else
                        MsgBox "Column is already on the left.", vbInformation
                    End If
                Case "right"
                    If selectedColIndex < .Columns.Count Then
                        Call MoveColumn(table, CopyTable, selectedColIndex, selectedColIndex + 1, textOnly, ignoreBorders)
                        MyDocument.Selection.ShapeRange.Delete
                        CopyTable.Select
                        CopyTable.table.Cell(selectedRowIndex, selectedColIndex + 1).Select
                    Else
                        MsgBox "Column is already on the right.", vbInformation
                    End If
                Case Else
                    MsgBox "Invalid direction. Please enter 'up', 'down', 'left', or 'right'.", vbCritical
            End Select
        End With
        
              
    Else
        MsgBox "No table selected or too many shapes selected. Select one table."
    End If
End Sub

Sub MoveRow(ByRef table As table, ByRef CopyTable, ByVal fromRow As Integer, ByVal toRow As Integer, textOnly As Boolean, ignoreBorders As Boolean)
    Dim i As Integer
    
    showWarning = False
    
    For i = 1 To table.Columns.Count

        table.Cell(fromRow, i).shape.TextFrame2.TextRange.Copy
        PauseForMilliseconds (50)
        CopyTable.table.Cell(toRow, i).shape.TextFrame2.TextRange.Paste
        table.Cell(toRow, i).shape.TextFrame2.TextRange.Copy
        PauseForMilliseconds (50)
        CopyTable.table.Cell(fromRow, i).shape.TextFrame2.TextRange.Paste
        PauseForMilliseconds (50)

        If textOnly = False Then
        
        If table.Cell(toRow, i).shape.Fill.Type = msoFillSolid Then

        CopyTable.table.Cell(fromRow, i).shape.Fill.Solid
        CopyTable.table.Cell(fromRow, i).shape.Fill.ForeColor.RGB = table.Cell(toRow, i).shape.Fill.ForeColor.RGB
        CopyTable.table.Cell(fromRow, i).shape.Fill.Transparency = table.Cell(toRow, i).shape.Fill.Transparency
        ElseIf table.Cell(toRow, i).shape.Fill.Type = -2 Then
        CopyTable.table.Cell(fromRow, i).shape.Fill.visible = False
        Else
        
        showWarning = True
                
        End If
        
        If table.Cell(fromRow, i).shape.Fill.Type = msoFillSolid Then
        CopyTable.table.Cell(toRow, i).shape.Fill.Solid
        CopyTable.table.Cell(toRow, i).shape.Fill.ForeColor.RGB = table.Cell(fromRow, i).shape.Fill.ForeColor.RGB
        CopyTable.table.Cell(toRow, i).shape.Fill.Transparency = table.Cell(fromRow, i).shape.Fill.Transparency
        ElseIf table.Cell(fromRow, i).shape.Fill.Type = -2 Then
        CopyTable.table.Cell(toRow, i).shape.Fill.visible = False
        Else
        showWarning = True
        End If

        CopyTable.table.Cell(fromRow, i).shape.TextFrame2.MarginLeft = table.Cell(toRow, i).shape.TextFrame2.MarginLeft
        CopyTable.table.Cell(fromRow, i).shape.TextFrame2.MarginRight = table.Cell(toRow, i).shape.TextFrame2.MarginRight
        CopyTable.table.Cell(fromRow, i).shape.TextFrame2.MarginTop = table.Cell(toRow, i).shape.TextFrame2.MarginTop
        CopyTable.table.Cell(fromRow, i).shape.TextFrame2.MarginBottom = table.Cell(toRow, i).shape.TextFrame2.MarginBottom

        If ignoreBorders = False Then
        Call CopyBorders(table, CopyTable, fromRow, i, toRow, i)
        End If
        
        End If
        
    Next i
    
        If showWarning = True Then
        MsgBox "This function only supports solid cell background fills (no gradients, textures or picture fills)"
        End If
    
End Sub


Sub MoveColumn(ByRef table As table, ByRef CopyTable, ByVal fromCol As Integer, ByVal toCol As Integer, textOnly As Boolean, ignoreBorders As Boolean)
    Dim i As Integer
    
    showWarning = False
    
    For i = 1 To table.Rows.Count

        table.Cell(i, fromCol).shape.TextFrame2.TextRange.Copy
        PauseForMilliseconds (50)
        CopyTable.table.Cell(i, toCol).shape.TextFrame2.TextRange.Paste
        table.Cell(i, toCol).shape.TextFrame2.TextRange.Copy
        PauseForMilliseconds (50)
        CopyTable.table.Cell(i, fromCol).shape.TextFrame2.TextRange.Paste
        PauseForMilliseconds (50)
        
        If textOnly = False Then
        
        If table.Cell(i, toCol).shape.Fill.Type = msoFillSolid Then
        CopyTable.table.Cell(i, fromCol).shape.Fill.Solid
        CopyTable.table.Cell(i, fromCol).shape.Fill.ForeColor.RGB = table.Cell(i, toCol).shape.Fill.ForeColor.RGB
        CopyTable.table.Cell(i, fromCol).shape.Fill.Transparency = table.Cell(i, toCol).shape.Fill.Transparency
        ElseIf table.Cell(i, toCol).shape.Fill.Type = -2 Then
        CopyTable.table.Cell(i, fromCol).shape.Fill.visible = False
        Else
        
        showWarning = True
                
        End If
                
        If table.Cell(i, fromCol).shape.Fill.Type = msoFillSolid Then
        CopyTable.table.Cell(i, toCol).shape.Fill.Solid
        CopyTable.table.Cell(i, toCol).shape.Fill.ForeColor.RGB = table.Cell(i, fromCol).shape.Fill.ForeColor.RGB
        CopyTable.table.Cell(i, toCol).shape.Fill.Transparency = table.Cell(i, fromCol).shape.Fill.Transparency
        ElseIf table.Cell(i, fromCol).shape.Fill.Type = -2 Then
        CopyTable.table.Cell(i, toCol).shape.Fill.visible = False
        Else
        
        showWarning = True
                
        End If
        
        CopyTable.table.Cell(i, fromCol).shape.TextFrame2.MarginLeft = table.Cell(i, toCol).shape.TextFrame2.MarginLeft
        CopyTable.table.Cell(i, fromCol).shape.TextFrame2.MarginRight = table.Cell(i, toCol).shape.TextFrame2.MarginRight
        CopyTable.table.Cell(i, fromCol).shape.TextFrame2.MarginTop = table.Cell(i, toCol).shape.TextFrame2.MarginTop
        CopyTable.table.Cell(i, fromCol).shape.TextFrame2.MarginBottom = table.Cell(i, toCol).shape.TextFrame2.MarginBottom
        
        If ignoreBorders = False Then
        Call CopyBorders(table, CopyTable, i, fromCol, i, toCol)
        End If
        
        End If
        
    Next i
    
        If showWarning = True Then
        MsgBox "This function only supports solid cell background fills (no gradients, textures or picture fills)"
        End If
        
End Sub


Sub CopyBorders(ByRef table As table, ByRef CopyTable, ByVal fromRow As Integer, ByVal fromCol As Integer, ByVal toRow As Integer, ByVal toCol As Integer)
    Dim borderIndex As Integer

    Set sourceCell = table.Cell(fromRow, fromCol)
    Set targetCell = CopyTable.table.Cell(toRow, toCol)
    
    If fromRow > toRow Then
    borderDirection = "up"
    ElseIf fromRow < toRow Then
    borderDirection = "down"
    ElseIf fromCol > toCol Then
    borderDirection = "left"
    ElseIf fromCol < toCol Then
    borderDirection = "right"
    End If
    

    For borderIndex = ppBorderTop To ppBorderRight
        On Error Resume Next

        With CopyTable.table.Cell(toRow, toCol).Borders(borderIndex)
            .Transparency = 1
            .Weight = 0
            .DashStyle = msoLineSolid
            .Style = msoLineNone
        End With
        
        With CopyTable.table.Cell(fromRow, fromCol).Borders(borderIndex)
            .Transparency = 1
            .Weight = 0
            .DashStyle = msoLineSolid
            .Style = msoLineNone
        End With

    Next borderIndex

    For borderIndex = ppBorderTop To ppBorderDiagonalUp
On Error Resume Next
            
        With CopyTable.table.Cell(fromRow, fromCol).Borders(borderIndex)
            
            If Not ((borderDirection = "right") And (borderIndex = ppBorderRight)) Then
            
            If Not ((borderDirection = "down") And (borderIndex = ppBorderBottom)) Then
            
            .Transparency = table.Cell(toRow, toCol).Borders(borderIndex).Transparency
            .ForeColor.RGB = table.Cell(toRow, toCol).Borders(borderIndex).ForeColor.RGB
            .Style = table.Cell(toRow, toCol).Borders(borderIndex).Style
            .DashStyle = table.Cell(toRow, toCol).Borders(borderIndex).DashStyle
            .Weight = table.Cell(toRow, toCol).Borders(borderIndex).Weight
            
            End If
            
            End If
            
        End With

        With CopyTable.table.Cell(toRow, toCol).Borders(borderIndex)
            .Transparency = table.Cell(fromRow, fromCol).Borders(borderIndex).Transparency
            .ForeColor.RGB = table.Cell(fromRow, fromCol).Borders(borderIndex).ForeColor.RGB
            .Style = table.Cell(fromRow, fromCol).Borders(borderIndex).Style
            .DashStyle = table.Cell(fromRow, fromCol).Borders(borderIndex).DashStyle
            .Weight = table.Cell(fromRow, fromCol).Borders(borderIndex).Weight
        End With
        
        

        On Error GoTo 0
    Next borderIndex
End Sub

Sub PauseForMilliseconds(milliseconds As Long)
    Dim startTime As Single
    startTime = Timer
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
    Loop
End Sub
