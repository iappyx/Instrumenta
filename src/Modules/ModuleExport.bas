Attribute VB_Name = "ModuleExport"
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

Sub EmailSelectedSlides()
    #If Mac Then
        MsgBox "This Function will not work on a Mac"
    #Else
    
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim OutlookApplication, OutlookMessage As Object
        Dim TemporaryPresentation, ThisPresentation As Presentation
        Dim PresentationFilename, EmailSubject As String
        Dim SlideLoop As Long
        Dim PresentationSlides As Slide
        Dim DotPosition As Integer
        
        Set ThisPresentation = ActivePresentation
        
        'Delete any previous export tags
        On Error Resume Next
        For Each PresentationSlides In ThisPresentation.Slides
            PresentationSlides.Tags.Delete ("INSTRUMENTA EXPORT")
        Next PresentationSlides
        
        'Strip extension from filename
        DotPosition = InStrRev(ThisPresentation.Name, ".")
        If DotPosition > 0 Then
            PresentationFilename = Left(ThisPresentation.Name, DotPosition - 1)
        Else
            PresentationFilename = ThisPresentation.Name
        End If
        
        'Set filename and e-mailsubject
        EmailSubject = PresentationFilename
        PresentationFilename = PresentationFilename & " (slide "
        
        ProgressForm.Show
        
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.Count
        
        SetProgress (SlideLoop / ActiveWindow.Selection.SlideRange.Count * 100)
        
            ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "INSTRUMENTA EXPORT", "YES"
            If SlideLoop <> ActiveWindow.Selection.SlideRange.Count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex
            End If
        Next SlideLoop
        
        ProgressForm.Hide
        
        PresentationFilename = PresentationFilename & ")"
        
        PresentationFilename = InputBox("Attachment file name:", "Send as e-mail", PresentationFilename)
        
        'Remove slides that where not selected for export
        ThisPresentation.SaveCopyAs Environ("TEMP") & "\" & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(Environ("TEMP") & "\" & PresentationFilename & ".pptx")
        For SlideLoop = TemporaryPresentation.Slides.Count To 1 Step -1
            If TemporaryPresentation.Slides(SlideLoop).Tags("INSTRUMENTA EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        TemporaryPresentation.Save
        TemporaryPresentation.Close
               
        On Error Resume Next
        Set OutlookApplication = GetObject(Class:="Outlook.Application")
        Err.Clear
        If OutlookApplication Is Nothing Then Set OutlookApplication = CreateObject(Class:="Outlook.Application")
        On Error GoTo 0
        Set OutlookMessage = OutlookApplication.CreateItem(0)
        
        On Error Resume Next
        With OutlookMessage
            .To = ""
            .CC = ""
            .Subject = EmailSubject
            .Body = ""
            .Attachments.Add Environ("TEMP") & "\" & PresentationFilename & ".pptx"
            .Display
        End With
        
        On Error GoTo 0
        
        'Clean temporary slides
        Set TemporaryPresentation = Presentations.Open(Environ("TEMP") & "\" & PresentationFilename & ".pptx")
        For SlideLoop = TemporaryPresentation.Slides.Count To 1 Step -1
            TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        TemporaryPresentation.Save
        TemporaryPresentation.Close
        
        Else
        MsgBox "No slides selected."
        End If
        
    #End If
    
End Sub

Sub EmailSelectedSlidesAsPDF()
    #If Mac Then
        MsgBox "This Function will not work on a Mac"
    #Else
        
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
        
        Dim OutlookApplication, OutlookMessage As Object
        Dim PresentationFilename, EmailSubject As String
        Dim SlideLoop As Long
        Dim DotPosition As Integer
        
        DotPosition = InStrRev(ActivePresentation.Name, ".")
        
        If DotPosition > 0 Then
            PresentationFilename = Left(ActivePresentation.Name, DotPosition - 1)
        Else
            PresentationFilename = ActivePresentation.Name
        End If
        
        'Set filename and e-mailsubject
        EmailSubject = PresentationFilename
        PresentationFilename = PresentationFilename & " (slide "
        
        ProgressForm.Show
        
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.Count
        
        SetProgress (SlideLoop / ActiveWindow.Selection.SlideRange.Count * 100)
        
            ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "INSTRUMENTA EXPORT", "YES"
            If SlideLoop <> ActiveWindow.Selection.SlideRange.Count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex
            End If
        Next SlideLoop
        
        ProgressForm.Hide
        
        PresentationFilename = PresentationFilename & ")"
        
        PresentationFilename = InputBox("Attachment file name:", "Send as e-mail", PresentationFilename)
      
        ActivePresentation.ExportAsFixedFormat Environ("TEMP") & "\" & PresentationFilename & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint, msoFalse, , , , , ppPrintSelection

        On Error Resume Next
        Set OutlookApplication = GetObject(Class:="Outlook.Application")
        Err.Clear
        If OutlookApplication Is Nothing Then Set OutlookApplication = CreateObject(Class:="Outlook.Application")
        On Error GoTo 0
        Set OutlookMessage = OutlookApplication.CreateItem(0)
        
        On Error Resume Next
        With OutlookMessage
            .To = ""
            .CC = ""
            .Subject = EmailSubject
            .Body = ""
            .Attachments.Add Environ("TEMP") & "\" & PresentationFilename & ".pdf"
            .Display
        End With
        
        On Error GoTo 0
        
        'Clean temporary PDF
        Dim FrontSlide As PrintRange
        ActivePresentation.PrintOptions.Ranges.ClearAll
        Set FrontSlide = ActivePresentation.PrintOptions.Ranges.Add(1, 1)
        
        ActivePresentation.ExportAsFixedFormat Environ("TEMP") & "\" & PresentationFilename & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint, msoFalse, , , , FrontSlide, ppPrintSlideRange

        Else
        MsgBox "No slides selected."
        End If
        
    #End If
    
End Sub
