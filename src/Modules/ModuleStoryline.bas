Attribute VB_Name = "ModuleStoryline"
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
