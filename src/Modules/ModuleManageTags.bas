Attribute VB_Name = "ModuleManageTags"
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

Global TypeOfTag    As String

Sub ShowFormSelectSlidesByTag()
    
    Dim SlideTags() As Variant
    Dim SlideTagIndex As Integer
    SlideTagIndex = 0
    
    SelectSlidesByTagForm.SlideTagComboBox.Clear
    SelectSlidesByTagForm.SelectSlidesByTagFrame.Enabled = False
    
    With SelectSlidesByTagForm.StampComboBox
        .Clear
        .AddItem "CONFIDENTIAL"
        .AddItem "DO NOT DISTRIBUTE"
        .AddItem "DRAFT"
        .AddItem "UPDATED"
        .AddItem "NEW"
        .AddItem "TO BE REMOVED"
        .AddItem "TO APPENDIX"
        .Value = "TO BE REMOVED"
    End With
    
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For TagCount = 1 To PresentationSlide.Tags.Count
            
            ReDim Preserve SlideTags(SlideTagIndex)
            SlideTags(SlideTagIndex) = PresentationSlide.Tags.Name(TagCount)
            SlideTagIndex = SlideTagIndex + 1
            
        Next TagCount
        
    Next
    
    SlideTags = RemoveDuplicates(SlideTags)
    
    For TagCount = 0 To UBound(SlideTags) - 1
        SelectSlidesByTagForm.SelectSlidesByTagFrame.Enabled = True
        SelectSlidesByTagForm.SlideTagComboBox.AddItem SlideTags(TagCount)
    Next TagCount
    
    SelectSlidesByTagForm.Show
    
End Sub

Sub PopulateSlideTagValueListbox()
    Dim SlideTagValues() As Variant
    Dim SlideTagValueIndex As Integer
    SlideTagValueIndex = 0
    
    SelectSlidesByTagForm.SlideTagValueListbox.Clear
    
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For TagCount = 1 To PresentationSlide.Tags.Count
            
            If PresentationSlide.Tags.Name(TagCount) = SelectSlidesByTagForm.SlideTagComboBox.Value Then
                ReDim Preserve SlideTagValues(SlideTagValueIndex)
                SlideTagValues(SlideTagValueIndex) = PresentationSlide.Tags.Value(TagCount)
                SlideTagValueIndex = SlideTagValueIndex + 1
            End If
            
        Next TagCount
        
    Next
    
    SlideTagValues = RemoveDuplicates(SlideTagValues)
    
    For TagCount = 0 To UBound(SlideTagValues) - 1
        SelectSlidesByTagForm.SlideTagValueListbox.AddItem
        SelectSlidesByTagForm.SlideTagValueListbox.List(TagCount, 0) = SlideTagValues(TagCount)
    Next TagCount
    
End Sub

Sub SelectSlidesByTag()
    
    Dim PresentationSlide As Slide
    Dim SlideSelection() As Variant
    Dim SlideIndex  As Integer
    Dim MatchFound  As Boolean
    
    SlideIndex = 0
    
    For SelectedCount = 0 To SelectSlidesByTagForm.SlideTagValueListbox.ListCount - 1
        If (SelectSlidesByTagForm.SlideTagValueListbox.Selected(SelectedCount) = True) Then
            
            For Each PresentationSlide In ActivePresentation.Slides
                
                For TagCount = 1 To PresentationSlide.Tags.Count
                    
                    If PresentationSlide.Tags.Name(TagCount) = SelectSlidesByTagForm.SlideTagComboBox.Value And PresentationSlide.Tags.Value(TagCount) = SelectSlidesByTagForm.SlideTagValueListbox.List(SelectedCount, 0) Then
                        
                        ReDim Preserve SlideSelection(SlideIndex)
                        SlideSelection(SlideIndex) = PresentationSlide.SlideIndex
                        SlideIndex = SlideIndex + 1
                        Exit For
                    End If
                    
                Next TagCount
                
            Next
            
        End If
        
    Next SelectedCount
    
    If SlideIndex > 0 Then
        SlideSelection = RemoveDuplicates(SlideSelection)
        Application.ActivePresentation.Slides.Range(SlideSelection).Select
    End If
    
    SelectSlidesByTagForm.Hide
    MsgBox Str(SlideIndex) & " slides selected with specified tag and value(s)."
    Unload SelectSlidesByTagForm
    
End Sub

Sub SelectSlidesByStamp(StampType As String)
    
    Dim PresentationSlide As Slide
    Dim SlideSelection() As Variant
    Dim SlideIndex  As Integer
    Dim MatchFound  As Boolean
    MatchFound = False
    
    SlideIndex = 0
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeCount = 1 To PresentationSlide.Shapes.Count
            
            For TagCount = 1 To PresentationSlide.Shapes(ShapeCount).Tags.Count
                
                If PresentationSlide.Shapes(ShapeCount).Tags.Name(TagCount) = "INSTRUMENTA STAMP" And PresentationSlide.Shapes(ShapeCount).Tags.Value(TagCount) = StampType Then
                    
                    ReDim Preserve SlideSelection(SlideIndex)
                    SlideSelection(SlideIndex) = PresentationSlide.SlideIndex
                    SlideIndex = SlideIndex + 1
                    MatchFound = True
                    Exit For
                End If
                
            Next TagCount
            
            If MatchFound = True Then
                MatchFound = False
                Exit For
            End If
            
        Next ShapeCount
        
    Next
    
    If SlideIndex > 0 Then
        Application.ActivePresentation.Slides.Range(SlideSelection).Select
    End If
    
    SelectSlidesByTagForm.Hide
    MsgBox Str(SlideIndex) & " slides selected with stamp " & StampType & "."
    Unload SelectSlidesByTagForm
    
End Sub

Sub ShowFormManageTags()
    Dim TotalCount  As Long
    TotalCount = 0
    
    ManageTagsForm.TagsListBox.Clear
    ManageTagsForm.TagsListBox.ColumnCount = 4
    ManageTagsForm.TagsListBox.ColumnWidths = "25;25;200;200"
    
    If Application.ActiveWindow.Selection.Type = ppSelectionShapes Then
        
        TypeOfTag = "shape"
        
        ManageTagsForm.FrameStandardTag.visible = False
        
        For ShapeCount = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
            
            For TagCount = 1 To Application.ActiveWindow.Selection.ShapeRange(ShapeCount).Tags.Count
                
                TotalCount = TotalCount + 1
                ManageTagsForm.TagsListBox.AddItem
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 0) = Str(ShapeCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 1) = Str(TagCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 2) = Application.ActiveWindow.Selection.ShapeRange(ShapeCount).Tags.Name(TagCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 3) = Application.ActiveWindow.Selection.ShapeRange(ShapeCount).Tags.Value(TagCount)
                
            Next
            
        Next
        ManageTagsForm.ShapeLabel.Caption = "Tags for selected shape(s):"
        ManageTagsForm.Show
        
    ElseIf Application.ActiveWindow.Selection.Type = ppSelectionSlides Then
        
        TypeOfTag = "slide"
        
        ManageTagsForm.FrameStandardTag.visible = True
        
        For SlideCount = 1 To Application.ActiveWindow.Selection.SlideRange.Count
            For TagCount = 1 To Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Count
                
                TotalCount = TotalCount + 1
                ManageTagsForm.TagsListBox.AddItem
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 0) = Str(SlideCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 1) = Str(TagCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 2) = Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Name(TagCount)
                ManageTagsForm.TagsListBox.List(TotalCount - 1, 3) = Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Value(TagCount)
                
            Next
            
        Next
        
        ManageTagsForm.ShapeLabel.Caption = "Tags for selected slide(s):"
        ManageTagsForm.Show
        
    Else
        MsgBox "No shapes or slides selected."
    End If
End Sub

Sub DeleteTag()
    
    If TypeOfTag = "slide" Then
        
        For SelectedCount = 0 To ManageTagsForm.TagsListBox.ListCount - 1
            If (ManageTagsForm.TagsListBox.Selected(SelectedCount) = True) Then
                
                Application.ActiveWindow.Selection.SlideRange(CLng(ManageTagsForm.TagsListBox.List(SelectedCount, 0))).Tags.Delete ManageTagsForm.TagsListBox.List(SelectedCount, 2)
                ManageTagsForm.Hide
                ShowFormManageTags
                
            End If
            
        Next SelectedCount
        
    ElseIf TypeOfTag = "shape" Then
        
        For SelectedCount = 0 To ManageTagsForm.TagsListBox.ListCount - 1
            If (ManageTagsForm.TagsListBox.Selected(SelectedCount) = True) Then
                
                Application.ActiveWindow.Selection.ShapeRange(CLng(ManageTagsForm.TagsListBox.List(SelectedCount, 0))).Tags.Delete ManageTagsForm.TagsListBox.List(SelectedCount, 2)
                ManageTagsForm.Hide
                ShowFormManageTags
                
            End If
            
        Next SelectedCount
        
    End If
    
End Sub

Sub DeleteAllTags()
    
    If TypeOfTag = "slide" Then
        
        If MsgBox("This will delete all tags above, are you sure?", vbYesNo) = vbNo Then Exit Sub
        
        For SelectedCount = 0 To ManageTagsForm.TagsListBox.ListCount - 1
            
            Application.ActiveWindow.Selection.SlideRange(CLng(ManageTagsForm.TagsListBox.List(SelectedCount, 0))).Tags.Delete ManageTagsForm.TagsListBox.List(SelectedCount, 2)
            
        Next SelectedCount
        ManageTagsForm.Hide
        ShowFormManageTags
        
    ElseIf TypeOfTag = "shape" Then
        
        If MsgBox("This will delete all tags above, are you sure?", vbYesNo) = vbNo Then Exit Sub
        
        For SelectedCount = 0 To ManageTagsForm.TagsListBox.ListCount - 1
            
            Application.ActiveWindow.Selection.ShapeRange(CLng(ManageTagsForm.TagsListBox.List(SelectedCount, 0))).Tags.Delete ManageTagsForm.TagsListBox.List(SelectedCount, 2)
            
        Next SelectedCount
        ManageTagsForm.Hide
        ShowFormManageTags
    End If
    
End Sub

Sub AddTag()
    
    If TypeOfTag = "slide" Then
        
        For SlideCount = 1 To Application.ActiveWindow.Selection.SlideRange.Count
            
            Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Add ManageTagsForm.AddTagIdTextBox.Value, ManageTagsForm.AddTagValueTextBox.Value
            
        Next SlideCount
        
        ManageTagsForm.AddTagIdTextBox.Value = ""
        ManageTagsForm.AddTagValueTextBox.Value = ""
        ManageTagsForm.Hide
        ShowFormManageTags
        
    ElseIf TypeOfTag = "shape" Then
        
        For ShapeCount = 1 To Application.ActiveWindow.Selection.ShapeRange.Count
            
            Application.ActiveWindow.Selection.ShapeRange(ShapeCount).Tags.Add ManageTagsForm.AddTagIdTextBox.Value, ManageTagsForm.AddTagValueTextBox.Value
            
        Next ShapeCount
        
        ManageTagsForm.AddTagIdTextBox.Value = ""
        ManageTagsForm.AddTagValueTextBox.Value = ""
        ManageTagsForm.Hide
        ShowFormManageTags
        
    End If
    
End Sub

Sub AddSpecialSlideTag(SpecialTagType As String)
    
    For SlideCount = 1 To Application.ActiveWindow.Selection.SlideRange.Count
        
        If SpecialTagType = "filename" Then
            
            Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Add "INSTRUMENTA ORIGINAL FILENAME", ActivePresentation.Name
            
        ElseIf SpecialTagType = "slidenum" Then
            
            Application.ActiveWindow.Selection.SlideRange(SlideCount).Tags.Add "INSTRUMENTA ORIGINAL SLIDENUM", Application.ActiveWindow.Selection.SlideRange(SlideCount).SlideNumber
            
        End If
        
    Next SlideCount
    
End Sub

Sub HideTagsOnSlide()
    
    Set MyDocument = Application.ActiveWindow
    
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "Tag") = 1 Then
                PresentationSlide.Shapes(ShapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub ShowTagsOnSlide()
    
    Set MyDocument = Application.ActiveWindow
    
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "Tag") = 1 Then
                PresentationSlide.Shapes(ShapeNumber).Delete
            End If
            
        Next
        
        For TagCount = 1 To PresentationSlide.Tags.Count
            
            RandomNumber = Round(Rnd() * 1000000, 0)
            
            Set TagBackground = PresentationSlide.Shapes.AddShape(msoShapeSnip2SameRectangle, 100, 100, 26, 150)
            
            With TagBackground
                .Line.visible = False
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Name = "TagBackground" + Str(RandomNumber)
                .Rotation = -90
            End With
            
            Set TagBackgroundInner = PresentationSlide.Shapes.AddShape(msoShapeOval, 43, 100 + 75 - 3, 6, 6)
            
            With TagBackgroundInner
                .Line.visible = False
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Name = "TagBackgroundInner" + Str(RandomNumber)
            End With
            
            Dim TagText As shape
            
            Set TagText = PresentationSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 53, 100 + 75 - 13, 135, 26)
            
            With TagText
                .TextFrame.AutoSize = ppAutoSizeNone
                .TextFrame.HorizontalAnchor = msoAnchorCenter
                .TextFrame.VerticalAnchor = msoAnchorMiddle
                .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextFrame.TextRange = PresentationSlide.Tags.Name(TagCount) + vbNewLine + PresentationSlide.Tags.Value(TagCount)
                .TextFrame.MarginBottom = 0
                .TextFrame.MarginTop = 0
                .TextFrame.MarginLeft = 0
                .TextFrame.MarginRight = 0
                
                .TextFrame.TextRange.Font.Bold = msoTrue
                .TextFrame.TextRange.Font.Name = "Arial"
                .TextFrame.TextRange.Font.Size = 7
                .Line.visible = False
                .Name = "TagText" + Str(RandomNumber)
            End With
            ActiveWindow.View.GotoSlide PresentationSlide.SlideNumber
            
            PresentationSlide.Shapes.Range(Array("TagBackground" + Str(RandomNumber), "TagBackgroundInner" + Str(RandomNumber), "TagText" + Str(RandomNumber))).Select
            CommandBars.ExecuteMso ("ShapesCombine")
            
            For Each shape In ActiveWindow.Selection.ShapeRange
                
                shape.Name = "Tag" + Str(RandomNumber)
                shape.Top = -95
                shape.left = 65 + (TagCount - 1) * (shape.Height + 5)
            Next
            
        Next TagCount
    Next
    ActiveWindow.Selection.Unselect
End Sub
