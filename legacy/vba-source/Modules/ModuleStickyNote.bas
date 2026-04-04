Attribute VB_Name = "ModuleStickyNote"
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

Sub GenerateStickyNote()
    
    Set MyDocument = Application.ActiveWindow
    
    Randomize
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.shapes.count
        
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Set StickyNote = MyDocument.Selection.SlideRange.shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.slideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
    
    With StickyNote
        .line.visible = False
        .Fill.ForeColor.RGB = GetSetting("Instrumenta", "StickyNotes", "StickyNotesColor", "49407")
        .Fill.Transparency = 0.1
        .name = "StickyNote" + Str(RandomNumber)
        
        With .TextFrame
            .marginBottom = 2
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 2
            .VerticalAnchor = msoAnchorTop
            .AutoSize = ppAutoSizeShapeToFitText
            
            With .textRange
                .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                .text = GetSetting("Instrumenta", "StickyNotes", "StickyNotesDefaultText", "Note")
                With .Font
                    .Size = 10
                    .color.RGB = RGB(0, 0, 0)
                End With
            End With
            
        End With
        
        .Tags.Add "INSTRUMENTA STICKYNOTE", NumberOfStickies
    End With
    
End Sub

Sub ConvertCommentsToStickyNotes()
    
    Set MyDocument = Application.ActiveWindow
    
    Randomize
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.shapes.count
        
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Dim CommentsCount As Long
    Dim RepliesCount As Long
    
    For CommentsCount = MyDocument.Selection.SlideRange.Comments.count To 1 Step -1
        
        Set StickyNote = MyDocument.Selection.SlideRange.shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.slideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
        
        With StickyNote
            .line.visible = False
            .Fill.ForeColor.RGB = GetSetting("Instrumenta", "StickyNotes", "StickyNotesColor", "49407")
            .Fill.Transparency = 0.1
            .name = "StickyNote" + Str(RandomNumber)
            .left = MyDocument.Selection.SlideRange.Comments(CommentsCount).left
            .Top = MyDocument.Selection.SlideRange.Comments(CommentsCount).Top
            .Tags.Add "INSTRUMENTA STICKYNOTE", NumberOfStickies
            
            With .TextFrame
                .marginBottom = 2
                .MarginLeft = 2
                .MarginRight = 2
                .MarginTop = 2
                .VerticalAnchor = msoAnchorTop
                .AutoSize = ppAutoSizeShapeToFitText
                
                With .textRange
                    .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                    .text = MyDocument.Selection.SlideRange.Comments(CommentsCount).Author & " (" & MyDocument.Selection.SlideRange.Comments(CommentsCount).AuthorInitials & "):" & vbNewLine & MyDocument.Selection.SlideRange.Comments(CommentsCount).text
                    With .Font
                        .Size = 10
                        .color.RGB = RGB(0, 0, 0)
                    End With
                    
                    For RepliesCount = MyDocument.Selection.SlideRange.Comments(CommentsCount).Replies.count To 1 Step -1
                        
                        .text = .text & vbNewLine & vbNewLine & MyDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).Author & " (" & MyDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).AuthorInitials & "):" & vbNewLine & MyDocument.Selection.SlideRange.Comments(CommentsCount).Replies(RepliesCount).text
                        
                    Next
                    
                End With
                
            End With
        End With
        
        MyDocument.Selection.SlideRange.Comments(CommentsCount).Delete
        NumberOfStickies = NumberOfStickies + 1
    Next
    
End Sub

Sub MoveStickyNotesOffSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.shapes.count
        
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            
            MyDocument.Selection.SlideRange.shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(MyDocument.Selection.SlideRange.shapes(ShapeNumber).Top)
            MyDocument.Selection.SlideRange.shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(MyDocument.Selection.SlideRange.shapes(ShapeNumber).left)
            
            
            With MyDocument.Selection.SlideRange.shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.slideWidth - .left - .width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.slideHeight - .Top - .height)
                             
            If .left <= ShapeRight And .left <= .Top And .left <= ShapeBottom Then
            
            .left = -5 - .width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .left Then
            
            .Top = -5 - .height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .left And ShapeRight <= .Top Then
            
            .left = 5 + Application.ActivePresentation.PageSetup.slideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.slideHeight
            
            End If
            
            End With
            End If
        
    Next
    
End Sub

Sub MoveStickyNotesOnSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = 1 To MyDocument.Selection.SlideRange.shapes.count
        On Error Resume Next
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            MyDocument.Selection.SlideRange.shapes(ShapeNumber).Top = CLng(MyDocument.Selection.SlideRange.shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            MyDocument.Selection.SlideRange.shapes(ShapeNumber).left = CLng(MyDocument.Selection.SlideRange.shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            
        End If
        On Error GoTo 0
    Next
    
End Sub

Sub DeleteStickyNotesOnSlide()
    Set MyDocument = Application.ActiveWindow
    
    For ShapeNumber = MyDocument.Selection.SlideRange.shapes.count To 1 Step -1
        
        If InStr(1, MyDocument.Selection.SlideRange.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            MyDocument.Selection.SlideRange.shapes(ShapeNumber).Delete
        End If
        
    Next
End Sub

Sub DeleteStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.shapes.count To 1 Step -1
            
            If InStr(1, PresentationSlide.shapes(ShapeNumber).name, "StickyNote") = 1 Then
                PresentationSlide.shapes(ShapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.shapes.count To 1 Step -1
            On Error Resume Next
            If InStr(1, PresentationSlide.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            PresentationSlide.shapes(ShapeNumber).Top = CLng(PresentationSlide.shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            PresentationSlide.shapes(ShapeNumber).left = CLng(PresentationSlide.shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            End If
            On Error GoTo 0
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOffAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.shapes.count To 1 Step -1
            
            If InStr(1, PresentationSlide.shapes(ShapeNumber).name, "StickyNote") = 1 Then
                
            PresentationSlide.shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(PresentationSlide.shapes(ShapeNumber).Top)
            PresentationSlide.shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(PresentationSlide.shapes(ShapeNumber).left)
            
            
            With PresentationSlide.shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.slideWidth - .left - .width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.slideHeight - .Top - .height)
                             
            If .left <= ShapeRight And .left <= .Top And .left <= ShapeBottom Then
            
            .left = -5 - .width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .left Then
            
            .Top = -5 - .height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .left And ShapeRight <= .Top Then
            
            .left = 5 + Application.ActivePresentation.PageSetup.slideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.slideHeight
            
            End If
            
            End With
            
            End If
            
        Next
        
    Next
    
End Sub

Sub ConvertAllCommentsToStickyNotes()
    
    Set MyDocument = Application.ActiveWindow
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    Randomize
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To PresentationSlide.shapes.count
        
        If InStr(1, PresentationSlide.shapes(ShapeNumber).name, "StickyNote") = 1 Then
            NumberOfStickies = NumberOfStickies + 1
        End If
        
    Next
    
    Dim CommentsCount As Long
    Dim RepliesCount As Long
    
    For CommentsCount = PresentationSlide.Comments.count To 1 Step -1
        
        Set StickyNote = PresentationSlide.shapes.AddShape(msoShapeRectangle, Application.ActivePresentation.PageSetup.slideWidth - (105 * (NumberOfStickies + 1)), 5, 100, 100)
        
        With StickyNote
            .line.visible = False
            .Fill.ForeColor.RGB = GetSetting("Instrumenta", "StickyNotes", "StickyNotesColor", "49407")
            .Fill.Transparency = 0.1
            .name = "StickyNote" + Str(RandomNumber)
            .left = PresentationSlide.Comments(CommentsCount).left
            .Top = PresentationSlide.Comments(CommentsCount).Top
            .Tags.Add "INSTRUMENTA STICKYNOTE", NumberOfStickies
            
            With .TextFrame
                .marginBottom = 2
                .MarginLeft = 2
                .MarginRight = 2
                .MarginTop = 2
                .VerticalAnchor = msoAnchorTop
                .AutoSize = ppAutoSizeShapeToFitText
                
                With .textRange
                    .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
                    .text = PresentationSlide.Comments(CommentsCount).Author & " (" & PresentationSlide.Comments(CommentsCount).AuthorInitials & "):" & vbNewLine & PresentationSlide.Comments(CommentsCount).text
                    With .Font
                        .Size = 10
                        .color.RGB = RGB(0, 0, 0)
                    End With
                    
                    For RepliesCount = PresentationSlide.Comments(CommentsCount).Replies.count To 1 Step -1
                        
                        .text = .text & vbNewLine & vbNewLine & PresentationSlide.Comments(CommentsCount).Replies(RepliesCount).Author & " (" & PresentationSlide.Comments(CommentsCount).Replies(RepliesCount).AuthorInitials & "):" & vbNewLine & PresentationSlide.Comments(CommentsCount).Replies(RepliesCount).text
                        
                    Next
                    
                End With
                
            End With
        End With
        
        PresentationSlide.Comments(CommentsCount).Delete
        NumberOfStickies = NumberOfStickies + 1
    Next
    
    Next
    
End Sub
