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

Sub ShowFormManageTags()
    Dim TotalCount  As Long
    TotalCount = 0
    
    ManageTagsForm.TagsListBox.Clear
    ManageTagsForm.TagsListBox.ColumnCount = 4
    ManageTagsForm.TagsListBox.ColumnWidths = "25;25;200;200"
    
    If Application.ActiveWindow.Selection.Type = ppSelectionShapes Then
        
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
       ' ManageTagsForm.ShapeLabel.Caption = "Tags for selected shape(s):"
        
        ManageTagsForm.Show
        
    ElseIf Application.ActiveWindow.Selection.Type = ppSelectionSlides Then
        
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
        MsgBox "No shapes Or slides selected."
    End If
End Sub
