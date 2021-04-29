Attribute VB_Name = "ModuleStoryline"
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

Sub CopyStorylineToClipboard()
    
    Dim SlideLoop   As Long
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    For Each PresentationSlide In ActivePresentation.Slides
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                StorylineText = StorylineText & SlidePlaceHolder.TextFrame.TextRange.Text & Chr(13)
                Exit For
            End If
        Next SlidePlaceHolder
    Next PresentationSlide
    
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    SlidePlaceHolder.TextFrame.TextRange.Text = StorylineText
    SlidePlaceHolder.TextFrame.TextRange.Copy
    SlidePlaceHolder.Delete
    
End Sub

Sub PasteStorylineInSelectedShape()
    
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    For Each PresentationSlide In ActivePresentation.Slides
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                StorylineText = StorylineText & SlidePlaceHolder.TextFrame.TextRange.Text & Chr(13)
                
                Exit For
            End If
        Next SlidePlaceHolder
    Next PresentationSlide
    
    Application.ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.Text = StorylineText
    
End Sub
