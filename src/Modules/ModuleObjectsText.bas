Attribute VB_Name = "ModuleObjectsText"
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

Sub ConvertTextToShapes()
    Dim shapeText       As shape
    Dim TempRectangle   As shape
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        For Each shapeText In ActiveWindow.Selection.ShapeRange
            
            If shapeText.HasTextFrame Then
                
                If shapeText.TextFrame2.HasText Then
                    
                    shapeText.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    
                    With shapeText.TextFrame2
                        Set TempRectangle = ActiveWindow.Selection.SlideRange.shapes.AddShape(msoShapeRectangle, shapeText.left, shapeText.Top, shapeText.width + .textRange.BoundWidth + .MarginRight, shapeText.height + .textRange.BoundHeight + .marginBottom)
                    End With
                    shapeText.Fill.visible = msoFalse
                    shapeText.line.visible = msoFalse
                    TempRectangle.Fill.visible = msoTrue
                    TempRectangle.line.visible = msoFalse
                    Set SlideShapeRange = ActiveWindow.Selection.SlideRange.shapes.Range(Array(shapeText.name, TempRectangle.name))
                    SlideShapeRange.Select
                    CommandBars.ExecuteMso ("ShapesIntersect")
                    
                End If
                
            End If
            
        Next shapeText
    End If
End Sub

Sub ObjectsTextToggleCase()
    
    Dim SlideShape  As shape
    Dim SlideTable  As table
    Dim SelectedTextRange As textRange
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                Set SelectedTextRange = SlideShape.TextFrame.textRange
                If Not SlideShape.Tags("CurrentCase") = "" Then
                    CurrentCase = Int(SlideShape.Tags("CurrentCase"))
                Else
                    CurrentCase = 0
                End If
                SlideShape.Tags.Add "CurrentCase", CurrentCase + 1
                SelectedTextRange.ChangeCase (1 + (CurrentCase + 1) Mod 4)
            End If
            
            If SlideShape.HasTable Then
                
                If Not SlideShape.Tags("CurrentCase") = "" Then
                    CurrentCase = Int(SlideShape.Tags("CurrentCase"))
                Else
                    CurrentCase = 0
                End If
                SlideShape.Tags.Add "CurrentCase", CurrentCase + 1
                
                Set SlideTable = SlideShape.table
                
                For i = 1 To SlideTable.rows.count
                    For j = 1 To SlideTable.Columns.count
                        Set SelectedTextRange = SlideTable.cell(i, j).shape.TextFrame.textRange
                        SelectedTextRange.ChangeCase (1 + (CurrentCase + 1) Mod 4)
                    Next j
                Next i
            End If
        Next SlideShape
    End If
    
End Sub

Sub ObjectsTextAddPeriods()
    
    Dim SlideTable  As table
    Dim SelectedTextRange As textRange
    Dim SlideShape  As shape
    
    Set MyDocument = Application.ActiveWindow
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each SlideShape In ActiveWindow.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                Set SelectedTextRange = SlideShape.TextFrame.textRange
                SelectedTextRange.AddPeriods
                
            End If
            
            If SlideShape.HasTable Then
                
                Set SlideTable = SlideShape.table
                
                For i = 1 To SlideTable.rows.count
                    For j = 1 To SlideTable.Columns.count
                        
                        Set SelectedTextRange = SlideTable.cell(i, j).shape.TextFrame.textRange
                        
                        SelectedTextRange.AddPeriods
                    Next j
                Next i
            End If
        Next SlideShape
    ElseIf sel.Type = ppSelectionText Then
        
        'sel.TextRange2.AddPeriods
        MsgBox "This Function only works reliably on shapes"
        
    End If
    
End Sub

Sub ObjectsTextRemovePeriods()
    
    Dim SlideTable  As table
    Dim SelectedTextRange As textRange
    Dim SlideShape  As shape
    
    Set MyDocument = Application.ActiveWindow
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each SlideShape In ActiveWindow.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                Set SelectedTextRange = SlideShape.TextFrame.textRange
                SelectedTextRange.RemovePeriods
                
            End If
            
            If SlideShape.HasTable Then
                
                Set SlideTable = SlideShape.table
                
                For i = 1 To SlideTable.rows.count
                    For j = 1 To SlideTable.Columns.count
                        
                        Set SelectedTextRange = SlideTable.cell(i, j).shape.TextFrame.textRange
                        
                        SelectedTextRange.RemovePeriods
                    Next j
                Next i
            End If
        Next SlideShape
    ElseIf sel.Type = ppSelectionText Then
        
        'sel.TextRange2.RemovePeriods
        MsgBox "This Function only works reliably on shapes"
        
    End If
    
End Sub

Sub ObjectsTextDeleteStrikethrough()
    Dim SlideShape      As shape
    Dim i, j, k         As Long
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
       
    ElseIf MyDocument.Selection.ShapeRange.count > 0 Then
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                If SlideShape.TextFrame.HasText Then
                    
                    For i = SlideShape.TextFrame2.textRange.Characters.count To 1 Step -1
                        If SlideShape.TextFrame2.textRange.Characters(i, 1).Font.Strikethrough = True Then
                            SlideShape.TextFrame2.textRange.Characters(i, 1).Delete
                        End If
                    Next i
                    
                End If
                
            ElseIf SlideShape.Type = msoGroup Then
                For Each GroupShape In SlideShape.GroupItems
                    If GroupShape.HasTextFrame Then
                        If GroupShape.TextFrame.HasText Then
                            
                            For j = GroupShape.TextFrame2.textRange.Characters.count To 1 Step -1
                                If GroupShape.TextFrame2.textRange.Characters(j, 1).Font.Strikethrough = True Then
                                    GroupShape.TextFrame2.textRange.Characters(j, 1).Delete
                                End If
                            Next j
                            
                        End If
                    End If
                    
                Next GroupShape
                
            Else
                
                If SlideShape.Type = msoTable Then
                    
                    Dim SlideTable As table
                    Set SlideTable = SlideShape.table
                    
                    For i = 1 To SlideTable.rows.count
                        For j = 1 To SlideTable.Columns.count
                            If SlideTable.cell(i, j).shape.HasTextFrame Then
                                If SlideTable.cell(i, j).shape.TextFrame.HasText Then
                                    
                                    For k = SlideTable.cell(i, j).shape.TextFrame2.textRange.Characters.count To 1 Step -1
                                        If SlideTable.cell(i, j).shape.TextFrame2.textRange.Characters(k, 1).Font.Strikethrough = True Then
                                            SlideTable.cell(i, j).shape.TextFrame2.textRange.Characters(k, 1).Delete
                                        End If
                                    Next k
                                    
                                End If
                            End If
                        Next j
                    Next i
                End If
            End If
            
        Next SlideShape
        
    Else
        
        MsgBox "No shapes or tables selected."
        
    End If
End Sub

Sub ObjectsTextColorBold(ColorAutomatic)
    Dim SlideShape      As shape
    Dim BoldColor       As Long
    BoldColor = -1
    Dim ShapeTextRange  As textRange
    Dim i, j, k         As Long
    
    Set MyDocument = Application.ActiveWindow
     
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    ElseIf MyDocument.Selection.ShapeRange.count > 0 Then
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                If SlideShape.TextFrame.HasText Then
                    
                    Set ShapeTextRange = SlideShape.TextFrame.textRange
                    
                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                        For i = 1 To ShapeTextRange.Characters.count
                            If ShapeTextRange.Characters(i, 1).Font.Bold = True Then
                                If BoldColor = -1 Then
                                    
                                    BoldColor = ShapeTextRange.Characters(i, 1).Font.color.RGB
                                    
                                    If ColorAutomatic = False Then
                                    BoldColor = ColorDialog(BoldColor)
                                    
                                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                    ShapeTextRange.Characters(i, 1).Font.color.RGB = BoldColor
                                    End If
                                    
                                    End If
                                    
                                Else
                                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                        ShapeTextRange.Characters(i, 1).Font.color = BoldColor
                                        
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                
            ElseIf SlideShape.Type = msoGroup Then
                For Each GroupShape In SlideShape.GroupItems
                    If GroupShape.HasTextFrame Then
                        If GroupShape.TextFrame.HasText Then
                            Set ShapeTextRange = GroupShape.TextFrame.textRange
                            If GroupShape.Fill.ForeColor.RGB <> BoldColor Then
                                For j = 1 To ShapeTextRange.Characters.count
                                    If ShapeTextRange.Characters(j, 1).Font.Bold = True Then
                                        If BoldColor = -1 Then
                                            BoldColor = ShapeTextRange.Characters(j, 1).Font.color.RGB
                                            
                                            If ColorAutomatic = False Then
                                            BoldColor = ColorDialog(BoldColor)
                                            
                                            If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                            ShapeTextRange.Characters(j, 1).Font.color.RGB = BoldColor
                                            End If
                                            
                                            End If
                                            
                                        Else
                                            If GroupShape.Fill.ForeColor.RGB <> BoldColor Then
                                                ShapeTextRange.Characters(j, 1).Font.color = BoldColor
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    End If
                Next GroupShape
                
            Else
                
                If SlideShape.Type = msoTable Then
                    
                    Dim SlideTable As table
                    Set SlideTable = SlideShape.table
                    
                    For i = 1 To SlideTable.rows.count
                        For j = 1 To SlideTable.Columns.count
                            If SlideTable.cell(i, j).shape.HasTextFrame Then
                                If SlideTable.cell(i, j).shape.TextFrame.HasText Then
                                    If SlideTable.cell(i, j).shape.Fill.ForeColor.RGB <> BoldColor Then
                                        Set ShapeTextRange = SlideTable.cell(i, j).shape.TextFrame.textRange
                                        For k = 1 To ShapeTextRange.Characters.count
                                            If ShapeTextRange.Characters(k, 1).Font.Bold = True Then
                                                If BoldColor = -1 Then
                                                    BoldColor = ShapeTextRange.Characters(k, 1).Font.color
                                                Else
                                                    If SlideTable.cell(i, j).shape.Fill.ForeColor.RGB <> BoldColor Then
                                                        ShapeTextRange.Characters(k, 1).Font.color = BoldColor
                                                    End If
                                                    
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                            End If
                        Next j
                    Next i
                End If
            End If
            
        Next SlideShape
        
    Else
        
        MsgBox "No shapes or tables selected."
        
    End If
End Sub


Sub ObjectsTextSplitByParagraph()

    Dim MyDocument As DocumentWindow
    Dim SlideShape As shape
    Dim DuplicateShape As shape
    Dim paraCount As Long
    Dim paraHeight As Single
    Dim i As Long

    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count > 1 Then
        
        MsgBox "Please Select one shape to split."
        
    ElseIf MyDocument.Selection.ShapeRange.count = 1 Then

    Set SlideShape = MyDocument.Selection.ShapeRange(1)

    If Not SlideShape.HasTextFrame Then
        MsgBox "The selected shape has no textframe."
        Exit Sub
    End If

    paraCount = SlideShape.TextFrame2.textRange.Paragraphs.count
    If paraCount = 0 Then Exit Sub

    paraHeight = SlideShape.height / paraCount

    For i = paraCount To 1 Step -1

        Set DuplicateShape = SlideShape.Duplicate(1)

        DuplicateShape.TextFrame2.textRange.text = ""

        SlideShape.TextFrame2.textRange.Paragraphs(i).Copy
        DoEvents

        DuplicateShape.TextFrame2.textRange.Characters(1, 0).Paste
        DoEvents

        DuplicateShape.Top = SlideShape.Top + paraHeight * (i - 1)
        DuplicateShape.left = SlideShape.left
        DuplicateShape.height = paraHeight

Dim tr As TextRange2
Dim p As TextRange2

Set tr = DuplicateShape.TextFrame2.textRange

If Not DuplicateShape.TextFrame2.HasText Then
    DuplicateShape.Delete
Else
    If tr.Paragraphs.count = 1 Then
        Set p = tr.Paragraphs(1)
        If p.Length = 1 And (p.text = vbCr Or p.text = vbLf) Then
            DuplicateShape.Delete
        End If
    End If
End If
        
    Next i
    
    SlideShape.Delete
    
    End If

End Sub

Sub MoveSelectedTextToNewElement()
    NewElementFromSelectedText True
End Sub

Sub CopySelectedTextToNewElement()
    NewElementFromSelectedText False
End Sub

Sub NewElementFromSelectedText(deleteOriginal As Boolean)

    Dim MyDocument As DocumentWindow
    Dim SlideShape As shape
    Dim DuplicateShape As shape
    Dim selRange As TextRange2

    Dim tbl As table
    Dim dupTbl As table
    Dim r As Long, c As Long
    Dim selRow As Long, selCol As Long
    Dim isTable As Boolean

    Dim cellRange As TextRange2
    Dim origCellRange As TextRange2
    Dim dupCellRange As TextRange2

    Dim selStart As Long
    Dim selLength As Long
    Dim relStart As Long
    Dim totalLength As Long

    Set MyDocument = Application.ActiveWindow

    If MyDocument.Selection.Type <> ppSelectionText Then Exit Sub
    If MyDocument.Selection.ShapeRange.count <> 1 Then Exit Sub

    Set SlideShape = MyDocument.Selection.ShapeRange(1)
    Set selRange = MyDocument.Selection.TextRange2
    If selRange.Length = 0 Then Exit Sub

    selStart = selRange.Start
    selLength = selRange.Length

    isTable = SlideShape.HasTable

    If isTable Then

        Set tbl = SlideShape.table

        Dim selFrame As TextFrame2
        Set selFrame = MyDocument.Selection.ShapeRange(1).table.cell(1, 1).shape.TextFrame2
        selRow = 0
        selCol = 0

        For r = 1 To tbl.rows.count
            For c = 1 To tbl.Columns.count

                Set cellRange = tbl.cell(r, c).shape.TextFrame2.textRange

                If tbl.cell(r, c).Selected Then
                    selRow = r
                    selCol = c
                    Exit For
                End If

            Next c
            If selRow > 0 Then Exit For
        Next r

        If selRow = 0 Then
            For r = 1 To tbl.rows.count
                For c = 1 To tbl.Columns.count
                    Set cellRange = tbl.cell(r, c).shape.TextFrame2.textRange
                    If selStart >= 1 And selStart <= cellRange.Length Then
                        If selLength <= cellRange.Length - selStart + 1 Then
                            Dim candidateText As String
                            candidateText = cellRange.Characters(selStart, selLength).text
                            If candidateText = selRange.text Then
                                selRow = r
                                selCol = c
                                Exit For
                            End If
                        End If
                    End If
                Next c
                If selRow > 0 Then Exit For
            Next r
        End If

        If selRow = 0 Then Exit Sub

        Set origCellRange = tbl.cell(selRow, selCol).shape.TextFrame2.textRange

        relStart = selStart

        Set DuplicateShape = SlideShape.Duplicate(1)
        Set dupTbl = DuplicateShape.table
        Set dupCellRange = dupTbl.cell(selRow, selCol).shape.TextFrame2.textRange

        totalLength = dupCellRange.Length

        If relStart + selLength <= totalLength Then
            dupCellRange.Characters(relStart + selLength, totalLength - (relStart + selLength) + 1).Delete
        End If

        If relStart > 1 Then
            dupCellRange.Characters(1, relStart - 1).Delete
        End If

        For r = 1 To dupTbl.rows.count
            For c = 1 To dupTbl.Columns.count
                If Not (r = selRow And c = selCol) Then
                    dupTbl.cell(r, c).shape.TextFrame2.textRange.text = ""
                End If
            Next c
        Next r

        Dim origCell As cell
        Dim dupCell As cell
        Set origCell = tbl.cell(selRow, selCol)
        Set dupCell = dupTbl.cell(selRow, selCol)

        Dim origFill As FillFormat
        Set origFill = origCell.shape.Fill
        With dupCell.shape.Fill
            Select Case origFill.Type
                Case msoFillSolid
                    .ForeColor.RGB = origFill.ForeColor.RGB
                    .Transparency = origFill.Transparency
                    .Solid
                Case msoFillBackground
                    .Background
                Case Else
                    .ForeColor.RGB = origFill.ForeColor.RGB
                    .Solid
            End Select
        End With

        Dim borderIdx As Long
        For borderIdx = ppBorderLeft To ppBorderDiagonalDown
            On Error Resume Next
            Dim origBorder As Object
            Dim dupBorder As Object
            Set origBorder = origCell.Borders(borderIdx)
            Set dupBorder = dupCell.Borders(borderIdx)
            If Not origBorder Is Nothing And Not dupBorder Is Nothing Then
                With dupBorder.line
                    .visible = origBorder.line.visible
                    If origBorder.line.visible = msoTrue Then
                        .ForeColor.RGB = origBorder.line.ForeColor.RGB
                        .Weight = origBorder.line.Weight
                        .DashStyle = origBorder.line.DashStyle
                    End If
                End With
            End If
            On Error GoTo 0
        Next borderIdx

        With dupCell.shape.TextFrame2
            .MarginLeft = origCell.shape.TextFrame2.MarginLeft
            .MarginRight = origCell.shape.TextFrame2.MarginRight
            .MarginTop = origCell.shape.TextFrame2.MarginTop
            .marginBottom = origCell.shape.TextFrame2.marginBottom
            .VerticalAnchor = origCell.shape.TextFrame2.VerticalAnchor
        End With

        If deleteOriginal Then
            origCellRange.Characters(relStart, selLength).Delete
        End If

        For r = dupTbl.rows.count To 1 Step -1
            If r <> selRow Then dupTbl.rows(r).Delete
        Next r

        For c = dupTbl.Columns.count To 1 Step -1
            If c <> selCol Then dupTbl.Columns(c).Delete
        Next c

        DuplicateShape.left = SlideShape.left + SlideShape.width + 10
        DuplicateShape.Top = SlideShape.Top

        Exit Sub
    End If

    Dim dupRange As TextRange2
    Dim startPos As Long
    Dim totalLen As Long

    Set DuplicateShape = SlideShape.Duplicate(1)
    Set dupRange = DuplicateShape.TextFrame2.textRange

    startPos = selRange.Start
    selLength = selRange.Length
    totalLen = dupRange.Length

    If startPos + selLength <= totalLen Then
        dupRange.Characters(startPos + selLength, totalLen - (startPos + selLength) + 1).Delete
    End If

    If startPos > 1 Then
        dupRange.Characters(1, startPos - 1).Delete
    End If

    If deleteOriginal Then
        SlideShape.TextFrame2.textRange.Characters(startPos, selLength).Delete
    End If

    DuplicateShape.left = SlideShape.left + SlideShape.width + 10
    DuplicateShape.Top = SlideShape.Top

End Sub

Sub ObjectsTextMerge()

    Dim MyDocument As DocumentWindow
    Dim SlideShape As shape
    Dim SlideShapeRange As ShapeRange
    Dim MergeShape As shape
    Dim tgt As TextRange2
    Dim src As TextRange2
    Dim i As Long

    Set MyDocument = Application.ActiveWindow

    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        Exit Sub
    End If

    If MyDocument.Selection.ShapeRange.count = 1 Then
        MsgBox "Please select more than one shape."
        Exit Sub
    End If

    Set SlideShapeRange = MyDocument.Selection.ShapeRange
    Set SlideShape = SlideShapeRange(1)

    If Not SlideShape.HasTextFrame Then
        MsgBox "The first selected shape has no textframe."
        Exit Sub
    End If

    Set tgt = SlideShape.TextFrame2.textRange

    If tgt.Length > 0 Then
        tgt.InsertAfter vbCrLf
    End If

    For i = 2 To SlideShapeRange.count

        Set MergeShape = SlideShapeRange(i)

        If MergeShape.HasTextFrame And MergeShape.TextFrame2.HasText Then

            Set src = MergeShape.TextFrame2.textRange

            If src.Length > 0 Then

                src.Copy
                DoEvents

                With tgt
                    .Characters(.Length + 1, 0).Paste
                End With
                DoEvents

                tgt.InsertAfter vbCrLf
            End If
        End If
    Next i

    For i = SlideShapeRange.count To 2 Step -1
        SlideShapeRange(i).Delete
    Next i

End Sub



Sub ObjectsTextInsertSpecialCharacter(SpecialCharacter As Long)
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        
        Application.ActiveWindow.Selection.textRange.InsertSymbol Application.ActiveWindow.Selection.textRange.Font.name, SpecialCharacter, MsoTriState.msoTrue
        
    End If
    
End Sub

Sub ObjectsIncreaseLineSpacing()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsLineSpacingLoop MyDocument.Selection.ChildShapeRange(i), 0.1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                
                ObjectsLineSpacingLoop MyDocument.Selection.ShapeRange(i), 0.1
                
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsDecreaseLineSpacing()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsLineSpacingLoop MyDocument.Selection.ChildShapeRange(i), -0.1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                
                ObjectsLineSpacingLoop MyDocument.Selection.ShapeRange(i), -0.1
                
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsLineSpacingLoop(SlideShape, LineSpacingChange)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsLineSpacingLoop SlideShapeChild, LineSpacingChange
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If Not ActiveWindow.Selection.Type = ppSelectionText Then
                
                Set textRange = SlideShape.TextFrame.textRange
                paragraphCount = textRange.Paragraphs.count
                
                If SlideShape.TextFrame2.AutoSize = msoAutoSizeTextToFitShape Then
                    
                    Dim fontSizes() As Single
                    Dim spaceWithin() As Single
                    
                    ReDim fontSizes(1 To textRange.Characters.count)
                    ReDim spaceWithin(1 To paragraphCount)
                    
                    For i = 1 To textRange.Characters.count
                        fontSizes(i) = textRange.Characters(i).Font.Size
                    Next i
                    
                    For i = 1 To paragraphCount
                        spaceWithin(i) = textRange.Paragraphs(i).ParagraphFormat.spaceWithin
                    Next i
                    
                    SlideShape.TextFrame2.AutoSize = msoAutoSizeNone
                    
                    For i = 1 To textRange.Characters.count
                        textRange.Characters(i).Font.Size = fontSizes(i)
                    Next i
                    
                    For i = 1 To paragraphCount
                        textRange.Paragraphs(i).ParagraphFormat.spaceWithin = spaceWithin(i)
                    Next i
                    
                    Erase fontSizes
                    Erase spaceWithin
                    
                End If
                
                If LineSpacingChange < 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            If .spaceWithin <= -LineSpacingChange Then
                                .spaceWithin = 0
                            Else
                                .spaceWithin = .spaceWithin + LineSpacingChange
                            End If
                        End With
                    Next i
                    
                ElseIf LineSpacingChange > 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            .spaceWithin = .spaceWithin + LineSpacingChange
                        End With
                    Next i
                End If
                
            Else
                
                Set textRange = ActiveWindow.Selection.textRange
                paragraphCount = textRange.Paragraphs.count
                
                If LineSpacingChange < 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            If .spaceWithin <= -LineSpacingChange Then
                                .spaceWithin = 0
                            Else
                                .spaceWithin = .spaceWithin + LineSpacingChange
                            End If
                        End With
                    Next i
                    
                ElseIf LineSpacingChange > 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            .spaceWithin = .spaceWithin + LineSpacingChange
                        End With
                    Next i
                End If
                
            End If
            
        End If
        
    End If
    
End Sub
Sub ObjectsRemoveText()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsRemoveTextLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsRemoveTextLoop MyDocument.Selection.ShapeRange(i)
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveTextLoop(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsRemoveTextLoop SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            SlideShape.TextFrame.textRange.text = ""
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveHyperlinks()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsRemoveHyperlinksLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
        Else
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsRemoveHyperlinksLoop MyDocument.Selection.ShapeRange(i)
            Next i
        End If
    End If
End Sub

Sub ObjectsRemoveHyperlinksLoop(SlideShape)
    If SlideShape.Type = msoGroup Then
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsRemoveHyperlinksLoop SlideShapeChild
        Next
    Else
        If SlideShape.HasTextFrame Then
                SlideShape.TextFrame.textRange.ActionSettings(ppMouseClick).Hyperlink.Delete
        End If
    End If
End Sub

Sub ObjectsSwapTextNoFormatting()
    
    Dim text1, text2 As String
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.ShapeRange.count = 2 Then
            
            If MyDocument.Selection.ShapeRange(1).HasTextFrame And MyDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                text1 = MyDocument.Selection.ShapeRange(1).TextFrame.textRange.text
                text2 = MyDocument.Selection.ShapeRange(2).TextFrame.textRange.text
                MyDocument.Selection.ShapeRange(1).TextFrame.textRange.text = text2
                MyDocument.Selection.ShapeRange(2).TextFrame.textRange.text = text1
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        ElseIf MyDocument.Selection.HasChildShapeRange Then
            
            
            If MyDocument.Selection.ChildShapeRange.count = 2 Then
                
                If MyDocument.Selection.ChildShapeRange(1).HasTextFrame And MyDocument.Selection.ChildShapeRange(2).HasTextFrame Then
                
                    text1 = MyDocument.Selection.ChildShapeRange(1).TextFrame.textRange.text
                    text2 = MyDocument.Selection.ChildShapeRange(2).TextFrame.textRange.text
                    MyDocument.Selection.ChildShapeRange(1).TextFrame.textRange.text = text2
                    MyDocument.Selection.ChildShapeRange(2).TextFrame.textRange.text = text1
                
                Else
            
                    MsgBox "Select two shapes that (can) have text."
            
                End If
                
            Else
                
                MsgBox "Select two shapes to swap their text."
            
            End If

        Else
            
            MsgBox "Select two shapes to swap their text."
            
        End If
        
    End If
    
End Sub

Sub ObjectsSwapText()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.ShapeRange.count = 2 Then
            
            If MyDocument.Selection.ShapeRange(1).HasTextFrame And MyDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                Dim SlidePlaceHolder As PowerPoint.shape
                Set SlidePlaceHolder = ActivePresentation.Slides(1).shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, width:=100, height:=100)
                
                MyDocument.Selection.ShapeRange(1).TextFrame.textRange.Cut
                SlidePlaceHolder.TextFrame.textRange.Paste
                
                MyDocument.Selection.ShapeRange(2).TextFrame.textRange.Cut
                MyDocument.Selection.ShapeRange(1).TextFrame.textRange.Paste
                
                SlidePlaceHolder.TextFrame.textRange.Cut
                MyDocument.Selection.ShapeRange(2).TextFrame.textRange.Paste
                
                SlidePlaceHolder.Delete
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        ElseIf MyDocument.Selection.HasChildShapeRange Then
            
            
            If MyDocument.Selection.ChildShapeRange.count = 2 Then
                
                If MyDocument.Selection.ChildShapeRange(1).HasTextFrame And MyDocument.Selection.ChildShapeRange(2).HasTextFrame Then
                               
                Dim SlidePlaceHolderChildShapeRange As PowerPoint.shape
                Set SlidePlaceHolderChildShapeRange = ActivePresentation.Slides(1).shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, width:=100, height:=100)
                
                MyDocument.Selection.ChildShapeRange(1).TextFrame.textRange.Cut
                SlidePlaceHolderChildShapeRange.TextFrame.textRange.Paste
                
                MyDocument.Selection.ChildShapeRange(2).TextFrame.textRange.Cut
                MyDocument.Selection.ChildShapeRange(1).TextFrame.textRange.Paste
                
                SlidePlaceHolderChildShapeRange.TextFrame.textRange.Cut
                MyDocument.Selection.ChildShapeRange(2).TextFrame.textRange.Paste
                
                SlidePlaceHolderChildShapeRange.Delete
                
                Else
            
                    MsgBox "Select two shapes that (can) have text."
            
                End If
                
            Else
                
                MsgBox "Select two shapes to swap their text."
            
            End If

        Else
            
            MsgBox "Select two shapes To swap their text."
            
        End If
        
    End If
    
End Sub

Sub ObjectsMarginsToZero()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), 0
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ShapeRange(i), 0
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsMarginsIncrease()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
    Dim ShapeMarginSetting As Double
    ShapeMarginSetting = CDbl(GetSetting("Instrumenta", "Shapes", "ShapeStepSizeMargin", "0" + GetDecimalSeperator() + "2"))
        
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), ShapeMarginSetting
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ShapeRange(i), ShapeMarginSetting
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsMarginsDecrease()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
    Dim ShapeMarginSetting As Double
    ShapeMarginSetting = CDbl(GetSetting("Instrumenta", "Shapes", "ShapeStepSizeMargin", "0" + GetDecimalSeperator() + "2"))
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), -ShapeMarginSetting
            Next i
            
        Else
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsMarginsLoop MyDocument.Selection.ShapeRange(i), -ShapeMarginSetting
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsMarginsLoop(SlideShape, MarginsChange)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsMarginsLoop SlideShapeChild, MarginsChange
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            With SlideShape.TextFrame
                
                If MarginsChange < 0 Then
                    
                    If .marginBottom >= -MarginsChange Then
                        .marginBottom = .marginBottom + MarginsChange
                    End If
                    If .MarginLeft >= -MarginsChange Then
                        .MarginLeft = .MarginLeft + MarginsChange
                    End If
                    If .MarginRight >= -MarginsChange Then
                        .MarginRight = .MarginRight + MarginsChange
                    End If
                    If .MarginTop >= -MarginsChange Then
                        .MarginTop = .MarginTop + MarginsChange
                    End If
                    
                ElseIf MarginsChange > 0 Then
                    
                    .marginBottom = .marginBottom + MarginsChange
                    .MarginLeft = .MarginLeft + MarginsChange
                    .MarginRight = .MarginRight + MarginsChange
                    .MarginTop = .MarginTop + MarginsChange
                    
                ElseIf MarginsChange = 0 Then
                    
                    .marginBottom = 0
                    .MarginLeft = 0
                    .MarginRight = 0
                    .MarginTop = 0
                    
                End If
                
            End With
            
        End If
        
    End If
    
End Sub

Sub ObjectsTextWordwrapToggle()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsTextWordwrapToggleLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsTextWordwrapToggleLoop MyDocument.Selection.ShapeRange(i)
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsTextWordwrapToggleLoop(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsTextWordwrapToggleLoop SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If SlideShape.TextFrame.WordWrap = True Then
                SlideShape.TextFrame.WordWrap = False
            Else
                SlideShape.TextFrame.WordWrap = True
            End If
            
        End If
        
    End If
    
End Sub

Sub ObjectsAutoSizeTextToFitShape()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 2
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ShapeRange(i), 2
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsAutoSizeShapeToFitText()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ShapeRange(i), 1
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsAutoSizeNone()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 0
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ShapeRange(i), 0
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsToggleAutoSize()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 5
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ShapeRange(i), 5
            Next i
            
        End If
        
    End If
End Sub

Sub ObjectsToggleAutoSizeLoop(SlideShape, AutoSizeNum)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsToggleAutoSizeLoop SlideShapeChild, AutoSizeNum
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If AutoSizeNum = 5 Then
                With SlideShape.TextFrame2
                    If .AutoSize = 0 Then
                        .AutoSize = 1
                    ElseIf .AutoSize = 1 Then
                        .AutoSize = 2
                    ElseIf .AutoSize = 2 Then
                        .AutoSize = 0
                    ElseIf .AutoSize = -2 Then
                        .AutoSize = 0
                    End If
                    
                End With
                
            Else
                
                SlideShape.TextFrame2.AutoSize = AutoSizeNum
                
            End If
            
        End If
        
    End If
    
End Sub


Sub ObjectsIncreaseLineSpacingBeforeAndAfter()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ChildShapeRange(i), 3
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ShapeRange(i), 3
                
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsDecreaseLineSpacingBeforeAndAfter()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.count
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ChildShapeRange(i), -3
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.count
                
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ShapeRange(i), -3
                
            Next i
            
        End If
        
    End If
    
End Sub

Sub ObjectsLineSpacingBeforeAndAfterLoop(SlideShape, LineSpacingChange)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsLineSpacingBeforeAndAfterLoop SlideShapeChild, LineSpacingChange
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If Not ActiveWindow.Selection.Type = ppSelectionText Then
                
                Set textRange = SlideShape.TextFrame.textRange
                paragraphCount = textRange.Paragraphs.count
                
                If SlideShape.TextFrame2.AutoSize = msoAutoSizeTextToFitShape Then
                    
                    Dim fontSizes() As Single
                    Dim spaceWithin() As Single
                    
                    ReDim fontSizes(1 To textRange.Characters.count)
                    ReDim spaceWithin(1 To paragraphCount)
                    
                    For i = 1 To textRange.Characters.count
                        fontSizes(i) = textRange.Characters(i).Font.Size
                    Next i
                    
                    For i = 1 To paragraphCount
                        spaceWithin(i) = textRange.Paragraphs(i).ParagraphFormat.spaceWithin
                    Next i
                    
                    SlideShape.TextFrame2.AutoSize = msoAutoSizeNone
                    
                    For i = 1 To textRange.Characters.count
                        textRange.Characters(i).Font.Size = fontSizes(i)
                    Next i
                    
                    For i = 1 To paragraphCount
                        textRange.Paragraphs(i).ParagraphFormat.spaceWithin = spaceWithin(i)
                    Next i
                    
                    Erase fontSizes
                    Erase spaceWithin
                    
                End If
                
                If LineSpacingChange < 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            If .SpaceBefore <= -LineSpacingChange Then
                                .SpaceBefore = 0
                                .SpaceAfter = 0
                            Else
                                .SpaceBefore = .SpaceBefore + LineSpacingChange
                                .SpaceAfter = .SpaceBefore
                            End If
                        End With
                    Next i
                    
                ElseIf LineSpacingChange > 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            .SpaceBefore = .SpaceBefore + LineSpacingChange
                            .SpaceAfter = .SpaceBefore
                        End With
                    Next i
                End If
                
            Else
                
                Set textRange = ActiveWindow.Selection.textRange
                paragraphCount = textRange.Paragraphs.count
                
                If LineSpacingChange < 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            If .SpaceBefore <= -LineSpacingChange Then
                                .SpaceBefore = 0
                                .SpaceAfter = 0
                            Else
                                .SpaceBefore = .SpaceBefore + LineSpacingChange
                                .SpaceAfter = .SpaceBefore
                            End If
                        End With
                    Next i
                    
                ElseIf LineSpacingChange > 0 Then
                    For i = 1 To paragraphCount
                        With textRange.Paragraphs(i).ParagraphFormat
                            .SpaceBefore = .SpaceBefore + LineSpacingChange
                            .SpaceAfter = .SpaceBefore
                        End With
                    Next i
                End If
                
            End If
            
        End If
        
    End If
    
End Sub
