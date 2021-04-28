Attribute VB_Name = "ModuleStickyNote"
Sub GenerateStickyNote()
    
    Set myDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Set StickyNote = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.SlideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
    
    With StickyNote
        .Line.Visible = False
        .Fill.ForeColor.RGB = RGB(255, 192, 0)
        .Fill.Transparency = 0.1
        .Name = "StickyNote" + Str(RandomNumber)
        
        With .TextFrame
            .MarginBottom = 2
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 2
            .VerticalAnchor = msoAnchorTop
            .AutoSize = ppAutoSizeShapeToFitText
            
            With .TextRange
                .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                .Text = "Note"
                With .Font
                    .Size = 10
                    .Color.RGB = RGB(0, 0, 0)
                End With
            End With
            
        End With
    End With
    
End Sub

Sub ConvertCommentsToStickyNotes()
    
    Set myDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Dim CommentsCount As Long
    Dim RepliesCount As Long
    
    For CommentsCount = myDocument.Selection.SlideRange.Comments.Count To 1 Step -1
        
        Set StickyNote = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.SlideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
        
        With StickyNote
            .Line.Visible = False
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            .Fill.Transparency = 0.1
            .Name = "StickyNote" + Str(RandomNumber)
            
            With .TextFrame
                .MarginBottom = 2
                .MarginLeft = 2
                .MarginRight = 2
                .MarginTop = 2
                .VerticalAnchor = msoAnchorTop
                .AutoSize = ppAutoSizeShapeToFitText
                
                With .TextRange
                    .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                    .Text = myDocument.Selection.SlideRange.Comments(CommentsCount).Author & " (" & myDocument.Selection.SlideRange.Comments(CommentsCount).AuthorInitials & "):" & vbNewLine & myDocument.Selection.SlideRange.Comments(CommentsCount).Text
                    With .Font
                        .Size = 10
                        .Color.RGB = RGB(0, 0, 0)
                    End With
                    
                    For RepliesCount = myDocument.Selection.SlideRange.Comments(CommentsCount).Replies.Count To 1 Step -1
                        
                        .Text = .Text & vbNewLine & vbNewLine & myDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).Author & " (" & myDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).AuthorInitials & "):" & vbNewLine & myDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).Text
                        
                    Next
                    
                End With
                
            End With
        End With
        
        myDocument.Selection.SlideRange.Comments(CommentsCount).Delete
        NumberOfStickies = NumberOfStickies + 1
    Next
    
End Sub

Sub MoveStickyNotesOffSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Top = -5 - myDocument.Selection.SlideRange.Shapes(shapeNumber).Height
        End If
        
    Next
    
End Sub

Sub MoveStickyNotesOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Top = 5
        End If
        
    Next
    
End Sub

Sub DeleteStickyNotesOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For shapeNumber = myDocument.Selection.SlideRange.Shapes.Count To 1 Step -1
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
            myDocument.Selection.SlideRange.Shapes(shapeNumber).Delete
        End If
        
    Next
End Sub

Sub DeleteStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
                PresentationSlide.Shapes(shapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
                PresentationSlide.Shapes(shapeNumber).Top = 5
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOffAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For shapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(shapeNumber).Name, "StickyNote") = 1 Then
                PresentationSlide.Shapes(shapeNumber).Top = -5 - PresentationSlide.Shapes(shapeNumber).Height
            End If
            
        Next
        
    Next
    
End Sub
