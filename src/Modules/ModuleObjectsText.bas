Attribute VB_Name = "ModuleObjectsText"
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

Sub ConvertTextToShapes()
    Dim ShapeText       As shape
    Dim TempRectangle   As shape
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        For Each ShapeText In ActiveWindow.Selection.ShapeRange
            
            If ShapeText.HasTextFrame Then
                
                If ShapeText.TextFrame2.HasText Then
                    
                    ShapeText.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    
                    With ShapeText.TextFrame2
                        Set TempRectangle = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, ShapeText.left, ShapeText.Top, ShapeText.Width + .TextRange.BoundWidth + .MarginRight, ShapeText.Height + .TextRange.BoundHeight + .MarginBottom)
                    End With
                    ShapeText.Fill.visible = msoFalse
                    ShapeText.Line.visible = msoFalse
                    TempRectangle.Fill.visible = msoTrue
                    TempRectangle.Line.visible = msoFalse
                    Set SlideShapeRange = ActiveWindow.Selection.SlideRange.Shapes.Range(Array(ShapeText.Name, TempRectangle.Name))
                    SlideShapeRange.Select
                    CommandBars.ExecuteMso ("ShapesIntersect")
                    
                End If
                
            End If
            
        Next ShapeText
    End If
End Sub

Sub ObjectsTextToggleCase()
    
    Dim SlideShape  As shape
    Dim SlideTable  As Table
    Dim SelectedTextRange As TextRange
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                Set SelectedTextRange = SlideShape.TextFrame.TextRange
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
                
                Set SlideTable = SlideShape.Table
                
                For i = 1 To SlideTable.Rows.Count
                    For j = 1 To SlideTable.Columns.Count
                        Set SelectedTextRange = SlideTable.Cell(i, j).shape.TextFrame.TextRange
                        SelectedTextRange.ChangeCase (1 + (CurrentCase + 1) Mod 4)
                    Next j
                Next i
            End If
        Next SlideShape
    End If
    
End Sub

Sub ObjectsTextAddPeriods()
    
    Dim SlideTable  As Table
    Dim SelectedTextRange As TextRange
    Dim SlideShape  As shape
    
    Set MyDocument = Application.ActiveWindow
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each SlideShape In ActiveWindow.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                Set SelectedTextRange = SlideShape.TextFrame.TextRange
                SelectedTextRange.AddPeriods
                
            End If
            
            If SlideShape.HasTable Then
                
                Set SlideTable = SlideShape.Table
                
                For i = 1 To SlideTable.Rows.Count
                    For j = 1 To SlideTable.Columns.Count
                        
                        Set SelectedTextRange = SlideTable.Cell(i, j).shape.TextFrame.TextRange
                        
                        SelectedTextRange.AddPeriods
                    Next j
                Next i
            End If
        Next SlideShape
    ElseIf Sel.Type = ppSelectionText Then
        
        'sel.TextRange2.AddPeriods
        MsgBox "This Function only works reliably on shapes"
        
    End If
    
End Sub

Sub ObjectsTextRemovePeriods()
    
    Dim SlideTable  As Table
    Dim SelectedTextRange As TextRange
    Dim SlideShape  As shape
    
    Set MyDocument = Application.ActiveWindow
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each SlideShape In ActiveWindow.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                Set SelectedTextRange = SlideShape.TextFrame.TextRange
                SelectedTextRange.RemovePeriods
                
            End If
            
            If SlideShape.HasTable Then
                
                Set SlideTable = SlideShape.Table
                
                For i = 1 To SlideTable.Rows.Count
                    For j = 1 To SlideTable.Columns.Count
                        
                        Set SelectedTextRange = SlideTable.Cell(i, j).shape.TextFrame.TextRange
                        
                        SelectedTextRange.RemovePeriods
                    Next j
                Next i
            End If
        Next SlideShape
    ElseIf Sel.Type = ppSelectionText Then
        
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
       
    ElseIf MyDocument.Selection.ShapeRange.Count > 0 Then
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                If SlideShape.TextFrame.HasText Then
                    
                    For i = SlideShape.TextFrame2.TextRange.Characters.Count To 1 Step -1
                        If SlideShape.TextFrame2.TextRange.Characters(i, 1).Font.Strikethrough = True Then
                            SlideShape.TextFrame2.TextRange.Characters(i, 1).Delete
                        End If
                    Next i
                    
                End If
                
            ElseIf SlideShape.Type = msoGroup Then
                For Each GroupShape In SlideShape.GroupItems
                    If GroupShape.HasTextFrame Then
                        If GroupShape.TextFrame.HasText Then
                            
                            For j = GroupShape.TextFrame2.TextRange.Characters.Count To 1 Step -1
                                If GroupShape.TextFrame2.TextRange.Characters(j, 1).Font.Strikethrough = True Then
                                    GroupShape.TextFrame2.TextRange.Characters(j, 1).Delete
                                End If
                            Next j
                            
                        End If
                    End If
                    
                Next GroupShape
                
            Else
                
                If SlideShape.Type = msoTable Then
                    
                    Dim SlideTable As Table
                    Set SlideTable = SlideShape.Table
                    
                    For i = 1 To SlideTable.Rows.Count
                        For j = 1 To SlideTable.Columns.Count
                            If SlideTable.Cell(i, j).shape.HasTextFrame Then
                                If SlideTable.Cell(i, j).shape.TextFrame.HasText Then
                                    
                                    For k = SlideTable.Cell(i, j).shape.TextFrame2.TextRange.Characters.Count To 1 Step -1
                                        If SlideTable.Cell(i, j).shape.TextFrame2.TextRange.Characters(k, 1).Font.Strikethrough = True Then
                                            SlideTable.Cell(i, j).shape.TextFrame2.TextRange.Characters(k, 1).Delete
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
    Dim ShapeTextRange  As TextRange
    Dim i, j, k         As Long
    
    Set MyDocument = Application.ActiveWindow
     
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    ElseIf MyDocument.Selection.ShapeRange.Count > 0 Then
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            
            If SlideShape.HasTextFrame Then
                
                If SlideShape.TextFrame.HasText Then
                    
                    Set ShapeTextRange = SlideShape.TextFrame.TextRange
                    
                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                        For i = 1 To ShapeTextRange.Characters.Count
                            If ShapeTextRange.Characters(i, 1).Font.Bold = True Then
                                If BoldColor = -1 Then
                                    
                                    BoldColor = ShapeTextRange.Characters(i, 1).Font.Color.RGB
                                    
                                    If ColorAutomatic = False Then
                                    BoldColor = ColorDialog(BoldColor)
                                    
                                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                    ShapeTextRange.Characters(i, 1).Font.Color.RGB = BoldColor
                                    End If
                                    
                                    End If
                                    
                                Else
                                    If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                        ShapeTextRange.Characters(i, 1).Font.Color = BoldColor
                                        
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
                            Set ShapeTextRange = GroupShape.TextFrame.TextRange
                            If GroupShape.Fill.ForeColor.RGB <> BoldColor Then
                                For j = 1 To ShapeTextRange.Characters.Count
                                    If ShapeTextRange.Characters(j, 1).Font.Bold = True Then
                                        If BoldColor = -1 Then
                                            BoldColor = ShapeTextRange.Characters(j, 1).Font.Color.RGB
                                            
                                            If ColorAutomatic = False Then
                                            BoldColor = ColorDialog(BoldColor)
                                            
                                            If SlideShape.Fill.ForeColor.RGB <> BoldColor Then
                                            ShapeTextRange.Characters(j, 1).Font.Color.RGB = BoldColor
                                            End If
                                            
                                            End If
                                            
                                        Else
                                            If GroupShape.Fill.ForeColor.RGB <> BoldColor Then
                                                ShapeTextRange.Characters(j, 1).Font.Color = BoldColor
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
                    
                    Dim SlideTable As Table
                    Set SlideTable = SlideShape.Table
                    
                    For i = 1 To SlideTable.Rows.Count
                        For j = 1 To SlideTable.Columns.Count
                            If SlideTable.Cell(i, j).shape.HasTextFrame Then
                                If SlideTable.Cell(i, j).shape.TextFrame.HasText Then
                                    If SlideTable.Cell(i, j).shape.Fill.ForeColor.RGB <> BoldColor Then
                                        Set ShapeTextRange = SlideTable.Cell(i, j).shape.TextFrame.TextRange
                                        For k = 1 To ShapeTextRange.Characters.Count
                                            If ShapeTextRange.Characters(k, 1).Font.Bold = True Then
                                                If BoldColor = -1 Then
                                                    BoldColor = ShapeTextRange.Characters(k, 1).Font.Color
                                                Else
                                                    If SlideTable.Cell(i, j).shape.Fill.ForeColor.RGB <> BoldColor Then
                                                        ShapeTextRange.Characters(k, 1).Font.Color = BoldColor
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
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.Count > 1 Then
        
        MsgBox "Please Select one shape to split."
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        
        Dim SlideShape As shape
        Dim ShapeHeight As Integer
        
        Set SlideShape = MyDocument.Selection.ShapeRange(1)
        
        If SlideShape.HasTextFrame Then
        
        ShapeHeight = SlideShape.Height / SlideShape.TextFrame.TextRange.Paragraphs.Count
        
        For i = SlideShape.TextFrame2.TextRange.Paragraphs.Count To 1 Step -1
            

            Set DuplicateShape = SlideShape.Duplicate
            
            
            SlideShape.TextFrame2.TextRange.Paragraphs(i).Copy
            DuplicateShape.TextFrame2.TextRange.Paste
            
            DuplicateShape.Top = SlideShape.Top + ShapeHeight * (i - 1)
            DuplicateShape.Height = ShapeHeight
            DuplicateShape.left = SlideShape.left
            
            If DuplicateShape.TextFrame2.TextRange.Text = "" Then
            
            DuplicateShape.Delete
            
            End If
            
        Next i
        
        SlideShape.Delete
        
        Else
        
            MsgBox "The selected shape has no textframe."
        
        End If
        
        
    End If
    
End Sub


Sub ObjectsTextMerge()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.Count = 1 Then
        MsgBox "Please select more than one shape."
        
    ElseIf MyDocument.Selection.ShapeRange.Count > 1 Then
        
        Dim SlideShape As shape
        Dim SlideShapeRange As ShapeRange

        Set SlideShapeRange = MyDocument.Selection.ShapeRange
        Set SlideShape = SlideShapeRange(1)
        
        If SlideShape.HasTextFrame Then
            
            SlideShapeRange(1).TextFrame.TextRange.InsertAfter vbCr
            
            For i = 2 To MyDocument.Selection.ShapeRange.Count
                Set MergeShape = SlideShapeRange(i)
                               
                If MergeShape.HasTextFrame Then
                    MergeShape.TextFrame2.TextRange.Copy
                    SlideShapeRange(1).TextFrame2.TextRange.InsertAfter(MergeShape.TextFrame2.TextRange).Paste
                    SlideShapeRange(1).Height = SlideShapeRange(1).Height + MergeShape.Height
                End If
                
            Next i
            
            For i = MyDocument.Selection.ShapeRange.Count To 2 Step -1
                
                If SlideShapeRange(i).HasTextFrame Then
                    SlideShapeRange(i).Delete
                End If
                
            Next i
            
        Else
            MsgBox "The first selected shape has no textframe."
            
        End If
        
    End If
    
End Sub


Sub ObjectsTextInsertSpecialCharacter(SpecialCharacter As Long)
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        
        Application.ActiveWindow.Selection.TextRange.InsertSymbol Application.ActiveWindow.Selection.TextRange.Font.Name, SpecialCharacter, MsoTriState.msoTrue
        
    End If
    
End Sub

Sub ObjectsIncreaseLineSpacing()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsLineSpacingLoop MyDocument.Selection.ChildShapeRange(i), 0.1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
                
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsLineSpacingLoop MyDocument.Selection.ChildShapeRange(i), -0.1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
                
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
            
            With SlideShape.TextFrame.TextRange.ParagraphFormat
                
                If LineSpacingChange < 0 Then
                    
                    If .SpaceWithin <= -LineSpacingChange Then
                        .SpaceWithin = 0
                    Else
                        .SpaceWithin = .SpaceWithin + LineSpacingChange
                    End If
                    
                ElseIf LineSpacingChange > 0 Then
                    
                    .SpaceWithin = .SpaceWithin + LineSpacingChange
                    
                End If
                
            End With
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveText()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsRemoveTextLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            SlideShape.TextFrame.TextRange.Text = ""
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveHyperlinks()
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsRemoveHyperlinksLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
        Else
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
                SlideShape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.Delete
        End If
    End If
End Sub

Sub ObjectsSwapTextNoFormatting()
    
    Dim text1, text2 As String
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        If MyDocument.Selection.ShapeRange.Count = 2 Then
            
            If MyDocument.Selection.ShapeRange(1).HasTextFrame And MyDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                text1 = MyDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text
                text2 = MyDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text
                MyDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
                MyDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        ElseIf MyDocument.Selection.HasChildShapeRange Then
            
            
            If MyDocument.Selection.ChildShapeRange.Count = 2 Then
                
                If MyDocument.Selection.ChildShapeRange(1).HasTextFrame And MyDocument.Selection.ChildShapeRange(2).HasTextFrame Then
                
                    text1 = MyDocument.Selection.ChildShapeRange(1).TextFrame.TextRange.Text
                    text2 = MyDocument.Selection.ChildShapeRange(2).TextFrame.TextRange.Text
                    MyDocument.Selection.ChildShapeRange(1).TextFrame.TextRange.Text = text2
                    MyDocument.Selection.ChildShapeRange(2).TextFrame.TextRange.Text = text1
                
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
        
        If MyDocument.Selection.ShapeRange.Count = 2 Then
            
            If MyDocument.Selection.ShapeRange(1).HasTextFrame And MyDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                Dim SlidePlaceHolder As PowerPoint.shape
                Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, Width:=100, Height:=100)
                
                MyDocument.Selection.ShapeRange(1).TextFrame.TextRange.Cut
                SlidePlaceHolder.TextFrame.TextRange.Paste
                
                MyDocument.Selection.ShapeRange(2).TextFrame.TextRange.Cut
                MyDocument.Selection.ShapeRange(1).TextFrame.TextRange.Paste
                
                SlidePlaceHolder.TextFrame.TextRange.Cut
                MyDocument.Selection.ShapeRange(2).TextFrame.TextRange.Paste
                
                SlidePlaceHolder.Delete
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        ElseIf MyDocument.Selection.HasChildShapeRange Then
            
            
            If MyDocument.Selection.ChildShapeRange.Count = 2 Then
                
                If MyDocument.Selection.ChildShapeRange(1).HasTextFrame And MyDocument.Selection.ChildShapeRange(2).HasTextFrame Then
                               
                Dim SlidePlaceHolderChildShapeRange As PowerPoint.shape
                Set SlidePlaceHolderChildShapeRange = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, Width:=100, Height:=100)
                
                MyDocument.Selection.ChildShapeRange(1).TextFrame.TextRange.Cut
                SlidePlaceHolderChildShapeRange.TextFrame.TextRange.Paste
                
                MyDocument.Selection.ChildShapeRange(2).TextFrame.TextRange.Cut
                MyDocument.Selection.ChildShapeRange(1).TextFrame.TextRange.Paste
                
                SlidePlaceHolderChildShapeRange.TextFrame.TextRange.Cut
                MyDocument.Selection.ChildShapeRange(2).TextFrame.TextRange.Paste
                
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), 0
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), ShapeMarginSetting
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsMarginsLoop MyDocument.Selection.ChildShapeRange(i), -ShapeMarginSetting
            Next i
            
        Else
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
                    
                    If .MarginBottom >= -MarginsChange Then
                        .MarginBottom = .MarginBottom + MarginsChange
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
                    
                    .MarginBottom = .MarginBottom + MarginsChange
                    .MarginLeft = .MarginLeft + MarginsChange
                    .MarginRight = .MarginRight + MarginsChange
                    .MarginTop = .MarginTop + MarginsChange
                    
                ElseIf MarginsChange = 0 Then
                    
                    .MarginBottom = 0
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsTextWordwrapToggleLoop MyDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 2
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 1
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 0
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop MyDocument.Selection.ChildShapeRange(i), 5
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ChildShapeRange(i), 3
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
                
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
            
            For i = 1 To MyDocument.Selection.ChildShapeRange.Count
                ObjectsLineSpacingBeforeAndAfterLoop MyDocument.Selection.ChildShapeRange(i), -3
            Next i
            
        Else
            
            For i = 1 To MyDocument.Selection.ShapeRange.Count
                
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
            
            With SlideShape.TextFrame.TextRange.ParagraphFormat
                
                If LineSpacingChange < 0 Then
                    
                    If .SpaceBefore <= -LineSpacingChange Then
                        .SpaceBefore = 0
            .SpaceAfter = 0
                    Else
                        .SpaceBefore = .SpaceBefore + LineSpacingChange
            .SpaceAfter = .SpaceBefore
                    End If
                    
                ElseIf LineSpacingChange > 0 Then
                    
                    .SpaceBefore = .SpaceBefore + LineSpacingChange
            .SpaceAfter = .SpaceBefore
                    
                End If
                
            End With
            
        End If
        
    End If
    
End Sub
