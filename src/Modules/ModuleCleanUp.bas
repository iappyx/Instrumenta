Attribute VB_Name = "ModuleCleanUp"
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

Sub CleanUpRemoveUnusedMasterSlides()
    Dim NumberOfDesigns, NumberOfCustomLayouts As Integer
    
    ProgressForm.Show
    
    On Error Resume Next
    
    DesignsCount = ActivePresentation.Designs.Count
    
    For NumberOfDesigns = ActivePresentation.Designs.Count To 1 Step -1
        
        SetProgress ((DesignsCount - NumberOfDesigns) / DesignsCount * 100)
        For NumberOfCustomLayouts = ActivePresentation.Designs(NumberOfDesigns).SlideMaster.CustomLayouts.Count To 1 Step -1
            ActivePresentation.Designs(NumberOfDesigns).SlideMaster.CustomLayouts(NumberOfCustomLayouts).Delete
        Next NumberOfCustomLayouts
        
        If ActivePresentation.Designs(NumberOfDesigns).SlideMaster.CustomLayouts.Count = 0 Then
            ActivePresentation.Designs(NumberOfDesigns).Delete
        End If
        
    Next NumberOfDesigns
    
    On Error GoTo 0
    
    ProgressForm.Hide
    Unload ProgressForm
    
End Sub

Public Function CallToSlideScopesForm() As String
    SlidesScopeForm.Show
    CallToSlideScopesForm = SlidesScopeForm.UserChoice
End Function

Sub CleanUpRemoveAnimationsFromAllSlides()
    Dim PresentationSlide As slide
    Dim AnimationCount As Long
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            ProgressForm.Show
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    For AnimationCount = PresentationSlide.TimeLine.MainSequence.Count To 1 Step -1
                        PresentationSlide.TimeLine.MainSequence.Item(AnimationCount).Delete
                    Next AnimationCount
                Next PresentationSlide
                
            End If
            
            ProgressForm.Hide
            Unload ProgressForm
            
        Case "all"
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                
                SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                
                For AnimationCount = PresentationSlide.TimeLine.MainSequence.Count To 1 Step -1
                    PresentationSlide.TimeLine.MainSequence.Item(AnimationCount).Delete
                Next AnimationCount
            Next PresentationSlide
            
            ProgressForm.Hide
            Unload ProgressForm
            
    End Select
    
End Sub

Sub CleanUpRemoveSpeakerNotesFromAllSlides()
    Dim PresentationSlide As slide
    Dim SlideShape  As PowerPoint.shape
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            ProgressForm.Show
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    For Each SlideShape In PresentationSlide.NotesPage.Shapes
                        If SlideShape.TextFrame.HasText Then
                            SlideShape.TextFrame.TextRange = ""
                        End If
                    Next
                    
                Next PresentationSlide
                
            End If
            
            ProgressForm.Hide
            Unload ProgressForm
            
        Case "all"
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                
                SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                
                For Each SlideShape In PresentationSlide.NotesPage.Shapes
                    If SlideShape.TextFrame.HasText Then
                        SlideShape.TextFrame.TextRange = ""
                    End If
                Next
            Next
            
            ProgressForm.Hide
            Unload ProgressForm
            
    End Select
    
End Sub

Sub CleanUpRemoveCommentsFromAllSlides()
    Dim PresentationSlide As slide
    Dim CommentsCount As Long
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            ProgressForm.Show
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    For CommentsCount = PresentationSlide.Comments.Count To 1 Step -1
                        PresentationSlide.Comments(1).Delete
                    Next
                    
                Next PresentationSlide
                
            End If
            
            ProgressForm.Hide
            Unload ProgressForm
            
        Case "all"
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                
                SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                
                For CommentsCount = PresentationSlide.Comments.Count To 1 Step -1
                    PresentationSlide.Comments(1).Delete
                Next
            Next
            
            ProgressForm.Hide
            Unload ProgressForm
            
    End Select
    
End Sub

Sub CleanUpRemoveSlideShowTransitionsFromAllSlides()
    Dim PresentationSlide As slide
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            ProgressForm.Show
            
            If TotalSelectedSlides > 0 Then
                SlideIndex = 0
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    SlideIndex = SlideIndex + 1
                    SetProgress (SlideIndex / TotalSelectedSlides * 100)
                    
                    PresentationSlide.SlideShowTransition.EntryEffect = 0
                    
                Next PresentationSlide
                
            End If
            
            ProgressForm.Hide
            Unload ProgressForm
            
        Case "all"
            
            ProgressForm.Show
            
            For Each PresentationSlide In ActivePresentation.Slides
                SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
                PresentationSlide.SlideShowTransition.EntryEffect = 0
            Next
            
            ProgressForm.Hide
            Unload ProgressForm
            
    End Select
    
End Sub

Sub CleanUpRemoveHiddenSlides()
    
    ProgressForm.Show
    
    NumberOfSlides = ActivePresentation.Slides.Count
    
    For SlideLoop = ActivePresentation.Slides.Count To 1 Step -1
        
        SetProgress ((NumberOfSlides - SlideLoop) / NumberOfSlides * 100)
        
        If ActivePresentation.Slides(SlideLoop).SlideShowTransition.Hidden = msoTrue Then
            
            ActivePresentation.Slides(SlideLoop).Delete
            
        End If
        
    Next
    
    ProgressForm.Hide
    Unload ProgressForm
    
End Sub

Sub CleanUpHideAndMoveSelectedSlides()
    NumberOfSlides = ActivePresentation.Slides.Count
    
    CurrentSlide = ActiveWindow.Selection.SlideRange(1).SlideIndex
    
    For i = ActiveWindow.Selection.SlideRange.Count To 1 Step -1
        ActivePresentation.Slides(ActiveWindow.Selection.SlideRange(i).SlideIndex).MoveTo (NumberOfSlides)
        ActivePresentation.Slides(NumberOfSlides).SlideShowTransition.Hidden = msoTrue
        NumberOfSlides = NumberOfSlides - 1
    Next i
    
    ActiveWindow.View.GotoSlide CurrentSlide
    
End Sub

Sub CleanUpAddSlideNumbers()
    Dim PresentationSlide As slide
    Dim SlideShape  As shape
    Dim hasSlideNumber As Boolean
    
    Select Case CallToSlideScopesForm()
        
        Case "cancel"
            
        Case "selected"
            
            TotalSelectedSlides = ActiveWindow.Selection.SlideRange.Count
            
            If TotalSelectedSlides > 0 Then
                
                For Each PresentationSlide In ActiveWindow.Selection.SlideRange
                    hasSlideNumber = False
                    
                    If PresentationSlide.SlideIndex = 1 Then GoTo NextSlideSelected
                    
                    For Each SlideShape In PresentationSlide.Shapes
                        If SlideShape.Type = msoPlaceholder Then
                            If SlideShape.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                                hasSlideNumber = True
                                Exit For
                            End If
                        End If
                    Next SlideShape
                    
                    If Not hasSlideNumber Then
                        PresentationSlide.HeadersFooters.SlideNumber.visible = True
                    End If
                    
NextSlideSelected:
                    
                Next PresentationSlide
                
            End If
            
        Case "all"
            
            For Each PresentationSlide In ActivePresentation.Slides
                hasSlideNumber = False
                
                If PresentationSlide.SlideIndex = 1 Then GoTo NextSlideAll
                
                For Each SlideShape In PresentationSlide.Shapes
                    If SlideShape.Type = msoPlaceholder Then
                        If SlideShape.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                            hasSlideNumber = True
                            Exit For
                        End If
                    End If
                Next SlideShape
                
                If Not hasSlideNumber Then
                    PresentationSlide.HeadersFooters.SlideNumber.visible = True
                End If
                
NextSlideAll:
            Next PresentationSlide
            
    End Select
    
End Sub
