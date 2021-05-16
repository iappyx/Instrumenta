Attribute VB_Name = "ModuleStickyNote"
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

Sub GenerateStickyNote()
    
    Set myDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
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
        
        .Tags.Add "INSTRUMENTA STICKYNOTE", NumberOfStickies
    End With
    
End Sub

Sub ConvertCommentsToStickyNotes()
    
    Set myDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim NumberOfStickies As Long
    NumberOfStickies = 0
    
    For ShapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
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
            .Left = myDocument.Selection.SlideRange.Comments(CommentsCount).Left
            .Top = myDocument.Selection.SlideRange.Comments(CommentsCount).Top
            .Tags.Add "INSTRUMENTA STICKYNOTE", NumberOfStickies
            
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
    
    For ShapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            
            myDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(myDocument.Selection.SlideRange.Shapes(ShapeNumber).Top)
            myDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(myDocument.Selection.SlideRange.Shapes(ShapeNumber).Left)
            
            
            With myDocument.Selection.SlideRange.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            End If
        
    Next
    
End Sub

Sub MoveStickyNotesOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For ShapeNumber = 1 To myDocument.Selection.SlideRange.Shapes.Count
        On Error Resume Next
        If InStr(1, myDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            myDocument.Selection.SlideRange.Shapes(ShapeNumber).Top = CLng(myDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            myDocument.Selection.SlideRange.Shapes(ShapeNumber).Left = CLng(myDocument.Selection.SlideRange.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            
        End If
        On Error GoTo 0
    Next
    
End Sub

Sub DeleteStickyNotesOnSlide()
    Set myDocument = Application.ActiveWindow
    
    For ShapeNumber = myDocument.Selection.SlideRange.Shapes.Count To 1 Step -1
        
        If InStr(1, myDocument.Selection.SlideRange.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            myDocument.Selection.SlideRange.Shapes(ShapeNumber).Delete
        End If
        
    Next
End Sub

Sub DeleteStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
                PresentationSlide.Shapes(ShapeNumber).Delete
            End If
            
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOnAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            On Error Resume Next
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
            PresentationSlide.Shapes(ShapeNumber).Top = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION TOP"))
            PresentationSlide.Shapes(ShapeNumber).Left = CLng(PresentationSlide.Shapes(ShapeNumber).Tags("INSTRUMENTA OLD POSITION LEFT"))
            End If
            On Error GoTo 0
        Next
        
    Next
    
End Sub

Sub MoveStickyNotesOffAllSlides()
    Dim PresentationSlide As Slide
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        For ShapeNumber = PresentationSlide.Shapes.Count To 1 Step -1
            
            If InStr(1, PresentationSlide.Shapes(ShapeNumber).Name, "StickyNote") = 1 Then
                
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION TOP", CStr(PresentationSlide.Shapes(ShapeNumber).Top)
            PresentationSlide.Shapes(ShapeNumber).Tags.Add "INSTRUMENTA OLD POSITION LEFT", CStr(PresentationSlide.Shapes(ShapeNumber).Left)
            
            
            With PresentationSlide.Shapes(ShapeNumber)
            ShapeRight = (Application.ActivePresentation.PageSetup.SlideWidth - .Left - .Width)
            ShapeBottom = (Application.ActivePresentation.PageSetup.SlideHeight - .Top - .Height)
                             
            If .Left <= ShapeRight And .Left <= .Top And .Left <= ShapeBottom Then
            
            .Left = -5 - .Width
            
            ElseIf .Top <= ShapeRight And .Top <= ShapeBottom And .Top <= .Left Then
            
            .Top = -5 - .Height
            
            ElseIf ShapeRight <= ShapeBottom And ShapeRight <= .Left And ShapeRight <= .Top Then
            
            .Left = 5 + Application.ActivePresentation.PageSetup.SlideWidth
            
            Else
            
            .Top = 5 + Application.ActivePresentation.PageSetup.SlideHeight
            
            End If
            
            End With
            
            End If
            
        Next
        
    Next
    
End Sub
