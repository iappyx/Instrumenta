Attribute VB_Name = "ModuleShapesAcrossSlides"
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

Sub DeleteTaggedShapes()
    Set MyDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For SlideCount = 1 To ActivePresentation.Slides.Count
                For Each shape In ActivePresentation.Slides(SlideCount).Shapes
                    
                    If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        shape.Delete
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub

Sub UpdateTaggedShapePositionAndDimensions()
    Set MyDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For SlideCount = 1 To ActivePresentation.Slides.Count
                For Each shape In ActivePresentation.Slides(SlideCount).Shapes
                    
                    If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        With shape
                            .Top = Application.ActiveWindow.Selection.ShapeRange.Top
                            .left = Application.ActiveWindow.Selection.ShapeRange.left
                            .Width = Application.ActiveWindow.Selection.ShapeRange.Width
                            .Height = Application.ActiveWindow.Selection.ShapeRange.Height
                            
                        End With
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub

Sub ShowFormCopyShapeToMultipleSlides()
    Set MyDocument = Application.ActiveWindow
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.Clear
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.ColumnCount = 3
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.ColumnWidths = "15;300;0"
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value = "NewShape" + Str(RandomNumber)
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Text = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
        End If
        
        Dim StorylineText As String
        Dim CurrentSlide As Long
        CurrentSlide = 0
        
        On Error Resume Next
        
        For SlideCount = 1 To ActivePresentation.Slides.Count
            
            If Not ActivePresentation.Slides(SlideCount).SlideNumber = Application.ActiveWindow.Selection.SlideRange.SlideNumber Then
                
                StorylineText = "Untitled"
                
                On Error Resume Next
                For Each SlidePlaceHolder In ActivePresentation.Slides(SlideCount).Shapes.Placeholders
                    
                    If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                        StorylineText = SlidePlaceHolder.TextFrame.TextRange.Text
                        Exit For
                    End If
                Next SlidePlaceHolder
                On Error GoTo 0
                
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.AddItem
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 0) = ActivePresentation.Slides(SlideCount).SlideNumber
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 1) = StorylineText
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 2) = ActivePresentation.Slides(SlideCount).SlideID
                
            Else
                CurrentSlide = 1
                
            End If
            
        Next SlideCount
        On Error GoTo 0
        
        CopyShapeToMultipleSlidesForm.Show
        
    Else
        MsgBox "No shapes selected."
    End If
End Sub

Sub CopyShapeToMultipleSlides()
    
    Dim shape       As shape
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim OverwriteExisting As Boolean
    Dim CrossSlideShapeId As String
    Dim SkipSlide   As Boolean
    
    OverwriteExisting = CopyShapeToMultipleSlidesForm.OptionExistingShapes1.Value
    CrossSlideShapeId = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value
    
    Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
    
    For SelectedCount = 0 To CopyShapeToMultipleSlidesForm.AllSlidesListBox.ListCount - 1
        If (CopyShapeToMultipleSlidesForm.AllSlidesListBox.Selected(SelectedCount) = True) Then
            
            SkipSlide = False
            
            For Each shape In ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SelectedCount))).Shapes
                
                If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                    
                    If OverwriteExisting = True Then
                        
                        shape.Delete
                        
                    Else
                        
                        SkipSlide = True
                        
                    End If
                    
                End If
                
            Next
            
            If SkipSlide = False Then
                Application.ActiveWindow.Selection.ShapeRange.Copy
                Set PastedShape = ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SelectedCount))).Shapes.Paste
                PastedShape.Name = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value + Str(RandomNumber)
                PastedShape.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
            End If
            
        End If
    Next SelectedCount
    
    CopyShapeToMultipleSlidesForm.Hide
    Unload CopyShapeToMultipleSlidesForm
    
End Sub
