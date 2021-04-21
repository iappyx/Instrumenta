Attribute VB_Name = "ModuleExport"
Sub EmailSelectedSlides()
    #If Mac Then
        MsgBox "This Function will Not work On a Mac"
    #Else
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
            PresentationSlides.Tags.Delete ("EXPORT")
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
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.Count
            ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "EXPORT", "YES"
            If SlideLoop <> ActiveWindow.Selection.SlideRange.Count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex
            End If
        Next SlideLoop
        PresentationFilename = PresentationFilename & ")"
        
        'Remove slides that where not selected for export
        ThisPresentation.SaveCopyAs Environ("TEMP") & "\" & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(Environ("TEMP") & "\" & PresentationFilename & ".pptx")
        For SlideLoop = TemporaryPresentation.Slides.Count To 1 Step -1
            Debug.Print TemporaryPresentation.Slides(SlideLoop).Tags("EXPORT")
            If TemporaryPresentation.Slides(SlideLoop).Tags("EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        TemporaryPresentation.Save
        TemporaryPresentation.Close
        
        Dim objOL   As Object
        Set objOL = CreateObject("Outlook.Application")
        
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
    #End If
    
End Sub

Sub EmailSelectedSlidesAsPDF()
    #If Mac Then
        MsgBox "This Function will Not work On a Mac"
    #Else
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
            PresentationSlides.Tags.Delete ("EXPORT")
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
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.Count
            ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "EXPORT", "YES"
            If SlideLoop <> ActiveWindow.Selection.SlideRange.Count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).SlideIndex
            End If
        Next SlideLoop
        PresentationFilename = PresentationFilename & ")"
        
        'Remove slides that where not selected for export
        ThisPresentation.SaveCopyAs Environ("TEMP") & "\" & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(Environ("TEMP") & "\" & PresentationFilename & ".pptx")
        For SlideLoop = TemporaryPresentation.Slides.Count To 1 Step -1
            Debug.Print TemporaryPresentation.Slides(SlideLoop).Tags("EXPORT")
            If TemporaryPresentation.Slides(SlideLoop).Tags("EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        TemporaryPresentation.Save
        
        ActivePresentation.ExportAsFixedFormat Environ("TEMP") & "\" & PresentationFilename & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint
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
            .Attachments.Add Environ("TEMP") & "\" & PresentationFilename & ".pdf"
            .Display
            
        End With
    #End If
    
End Sub
