VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertSlideLibrarySlide 
   Caption         =   "Insert slide from slide library"
   ClientHeight    =   6797
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   14028
   OleObjectBlob   =   "InsertSlideLibrarySlide.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertSlideLibrarySlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_Change()
    
    Dim TempPath     As String
    If ListBox1.Enabled = True Then
        #If Mac Then
            TempPath = MacScript("return posix path of (path to temporary items) as string")
        #Else
            
            TempPath = Environ("TEMP") & "\"
            
        #End If
    
    
        InsertSlideLibrarySlide.PreviewImage.Picture = LoadPicture(TempPath & "tmp.Slide" & ListBox1.ListIndex + 1 & ".jpg")
        InsertSlideLibrarySlide.Repaint
    End If
    
End Sub

Private Sub CommandButton1_Click()
    
    Dim LibraryPresentation       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim TempPath     As String
    Set MyDocument = Application.ActiveWindow
    Set Testdocument = Application.ActiveWindow.Presentation
    
    'Set MyApplication = Application
    
    #If Mac Then
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    LibraryPresentation.Slides.Item(ListBox1.ListIndex + 1).Copy
    
    LibraryPresentation.Close
    Set LibraryPresentation = Nothing
    
    Testdocument.Windows(1).Activate
    Testdocument.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")
    
    
    #If Mac Then
        
        TempPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        
        TempPath = Environ("TEMP") & "\"
        
    #End If
    
    If InsertSlideLibrarySlide.ListBox1.Enabled = True Then
        
        For i = 1 To InsertSlideLibrarySlide.ListBox1.ListCount
            
            Kill (TempPath & "tmp.Slide" & i & ".jpg")

        Next i
        
    End If
    

    
    InsertSlideLibrarySlide.ListBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    
    If InsertSlideLibrarySlide.ListBox1.Enabled = True Then
        
        Dim TempPath As String
        
        #If Mac Then
            

            TempPath = MacScript("return posix path of (path to temporary items) as string")

            Set LibraryPresentation = Nothing
        #Else
            
            TempPath = Environ("TEMP") & "\"
            
        #End If

        
        For i = 1 To InsertSlideLibrarySlide.ListBox1.ListCount
            
            Kill (TempPath & "tmp.Slide" & i & ".jpg")
            
        Next i
        
    End If
    
    InsertSlideLibrarySlide.ListBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Private Sub CommandButton3_Click()
    Dim LibraryPresentation       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim TempPath     As String
    Set MyDocument = Application.ActiveWindow
    
    #If Mac Then
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    LibraryPresentation.Slides.Item(ListBox1.ListIndex + 1).Copy
    
    LibraryPresentation.Close
    Set LibraryPresentation = Nothing
    
    
    MyDocument.Presentation.Slides.Paste
    
    'clean up
    #If Mac Then
        TempPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        TempPath = Environ("TEMP") & "\"
    #End If
    
    If InsertSlideLibrarySlide.ListBox1.Enabled = True Then
        
        For i = 1 To InsertSlideLibrarySlide.ListBox1.ListCount
            Kill (TempPath & "tmp.Slide" & i & ".jpg")
        Next i
        
    End If

    
    InsertSlideLibrarySlide.ListBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
    
    
    If GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "") = "" Then
    InsertSlideLibrarySlide.Hide
    MsgBox "No slide library file set. Please set the file in Instrumenta settings."
    
    SettingsForm.Show
    
    Else
    
   
    'Needed to enable OLE automation for this
    
    Dim LibraryPresentation       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim TempPath     As String
    
    #If Mac Then

    
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    #If Mac Then
        TempPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        TempPath = Environ("TEMP") & "\"
    #End If
    
    NumberOfSlides = LibraryPresentation.Slides.Count
    InsertSlideLibrarySlide.ListBox1.Clear
    
    For Each PresentationSlide In LibraryPresentation.Slides
        
        PresentationSlide.Export TempPath & "tmp.Slide" & PresentationSlide.SlideNumber & ".jpg", "JPG"
        
        SlideTitle = ""
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                SlideTitle = SlidePlaceHolder.TextFrame.TextRange.Text
                Exit For
            End If
        Next SlidePlaceHolder
        
        If SlideTitle = "" Then SlideTitle = PresentationSlide.Name
        InsertSlideLibrarySlide.ListBox1.AddItem SlideTitle
        
    Next
    
    LibraryPresentation.Close
    Set LibraryPresentation = Nothing
    
    InsertSlideLibrarySlide.PreviewImage.Picture = LoadPicture(TempPath & "tmp.Slide1.jpg")
    InsertSlideLibrarySlide.ListBox1.Enabled = True
    InsertSlideLibrarySlide.ListBox1.ListIndex = 0
    InsertSlideLibrarySlide.Repaint
    
    End If
    
    
End Sub


