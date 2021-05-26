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


Sub CopySlideNotesToClipboard(ExportToWord As Boolean)
    
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    Dim PlaceHolderTextRange As TextRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.TextRange
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        If PresentationSlide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
            
            PresentationSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Copy
            
            PlaceHolderTextRange.Characters(0).InsertAfter Chr(13) & Chr(13) & "[Slide " & Str(PresentationSlide.SlideNumber) & "]" & Chr(13)
            PlaceHolderTextRange.Characters(0).Paste
            
        End If
        
    Next PresentationSlide
    
    ProgressForm.Hide

    SlidePlaceHolder.TextFrame.TextRange.Copy
    SlidePlaceHolder.Delete
    
    If ExportToWord = True Then
    
        #If Mac Then
        MsgBox "This Function will not work on a Mac. Slide notes are copied to clipboard."
        #Else
    
            Dim WordApplication, WordDocument As Object
            
            On Error Resume Next
            Set WordApplication = GetObject(Class:="Word.Application")
            Err.Clear
            
            If WordApplication Is Nothing Then Set WordApplication = CreateObject(Class:="Word.Application")
            On Error GoTo 0
            
            WordApplication.Visible = True
            Set WordDocument = WordApplication.Documents.Add
            
            With WordApplication
                .Selection.PasteAndFormat wdPasteDefault
            End With
    
        #End If
    
    Else
        MsgBox "Slide notes copied to clipboard."
    End If
    
End Sub


Sub CopyStorylineToClipboard(ExportToWord As Boolean)
    
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
     
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                StorylineText = StorylineText & SlidePlaceHolder.TextFrame.TextRange.text & Chr(13)
                Exit For
            End If
        Next SlidePlaceHolder
    Next PresentationSlide
    
    ProgressForm.Hide
    
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    SlidePlaceHolder.TextFrame.TextRange.text = StorylineText
    SlidePlaceHolder.TextFrame.TextRange.Copy
    SlidePlaceHolder.Delete
    
    If Not StorylineText = "" Then
    If ExportToWord = True Then
    
        #If Mac Then
        MsgBox "This Function will not work on a Mac. Storyline is copied to clipboard."
        #Else
    
            Dim WordApplication, WordDocument As Object
            
            On Error Resume Next
            Set WordApplication = GetObject(Class:="Word.Application")
            Err.Clear
            
            If WordApplication Is Nothing Then Set WordApplication = CreateObject(Class:="Word.Application")
            On Error GoTo 0
            
            WordApplication.Visible = True
            Set WordDocument = WordApplication.Documents.Add
            
            With WordApplication
                
                    .Selection.PasteAndFormat wdPasteDefault
        
            End With
    
        #End If
    
    Else
        MsgBox "Storyline copied to clipboard."
    End If
    
    Else
        MsgBox "Storyline is empty"
    End If
    
End Sub

Sub PasteStorylineInSelectedShape()
    
    Set myDocument = Application.ActiveWindow
    If Not myDocument.Selection.Type = ppSelectionShapes Then
    MsgBox "Please select a shape."
    Else
    
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
    
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                StorylineText = StorylineText & SlidePlaceHolder.TextFrame.TextRange.text & Chr(13)
                
                Exit For
            End If
        Next SlidePlaceHolder
    Next PresentationSlide
    
    ProgressForm.Hide
    
    Application.ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.text = StorylineText
    
    End If
    
End Sub
