Attribute VB_Name = "ModuleCleanUp"
Sub CleanUpRemoveAnimationsFromAllSlides()
    Dim PresentationSlide As Slide
    Dim AnimationCount As Long
    
    For Each PresentationSlide In ActivePresentation.Slides
        For AnimationCount = PresentationSlide.TimeLine.MainSequence.Count To 1 Step -1
            PresentationSlide.TimeLine.MainSequence.Item(AnimationCount).Delete
        Next AnimationCount
    Next PresentationSlide
    
End Sub

Sub CleanUpRemoveSpeakerNotesFromAllSlides()
    Dim PresentationSlide As Slide
    Dim SlideShape  As PowerPoint.Shape
    
    For Each PresentationSlide In ActivePresentation.Slides
        For Each SlideShape In PresentationSlide.NotesPage.Shapes
            If SlideShape.TextFrame.HasText Then
                SlideShape.TextFrame.TextRange = ""
            End If
        Next
    Next
End Sub

Sub CleanUpRemoveCommentsFromAllSlides()
    Dim PresentationSlide As Slide
    Dim CommentsCount As Long
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For CommentsCount = PresentationSlide.Comments.Count To 1 Step -1
            PresentationSlide.Comments(1).Delete
        Next
    Next
End Sub

Sub CleanUpRemoveSlideShowTransitionsFromAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        PresentationSlide.SlideShowTransition.EntryEffect = 0
    Next
End Sub
