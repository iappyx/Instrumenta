VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertSlideLibrarySlide 
   Caption         =   "Insert slide from slide library"
   ClientHeight    =   9373.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   14434
   OleObjectBlob   =   "InsertSlideLibrarySlide.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertSlideLibrarySlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    
    Dim strPath     As String
    If ComboBox1.Enabled = True Then
        #If Mac Then
            strPath = MacScript("return posix path of (path to temporary items) as string")
        #Else
            
            strPath = Environ("TEMP") & "\"
            
        #End If
    
    
        InsertSlideLibrarySlide.PreviewImage.Picture = LoadPicture(strPath & "tmp.Slide" & ComboBox1.ListIndex + 1 & ".jpg")
        InsertSlideLibrarySlide.Repaint
    End If
    
End Sub

Private Sub CommandButton1_Click()
    
    Dim oPres       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim strPath     As String
    Set MyDocument = Application.ActiveWindow
    Set MyApplication = Application
    
    #If Mac Then
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    oPres.Slides.Item(ComboBox1.ListIndex + 1).Copy
    
    oPres.Close
    Set oPres = Nothing
    
    MyApplication.CommandBars.ExecuteMso ("PasteSourceFormatting")
    
    
    #If Mac Then
        
        strPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        
        strPath = Environ("TEMP") & "\"
        
    #End If
    
    If InsertSlideLibrarySlide.ComboBox1.Enabled = True Then
        
        For i = 1 To InsertSlideLibrarySlide.ComboBox1.ListCount
            
            Kill (strPath & "tmp.Slide" & i & ".jpg")

        Next i
        
    End If
    

    
    InsertSlideLibrarySlide.ComboBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
    
End Sub

Private Sub CommandButton2_Click()
    
    If InsertSlideLibrarySlide.ComboBox1.Enabled = True Then
        
        Dim strPath As String
        
        #If Mac Then
            

            strPath = MacScript("return posix path of (path to temporary items) as string")

            Set oPres = Nothing
        #Else
            
            strPath = Environ("TEMP") & "\"
            
        #End If

        
        For i = 1 To InsertSlideLibrarySlide.ComboBox1.ListCount
            
            Kill (strPath & "tmp.Slide" & i & ".jpg")
            
        Next i
        
    End If
    
    InsertSlideLibrarySlide.ComboBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
    
End Sub

Private Sub CommandButton3_Click()
    Dim oPres       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim strPath     As String
    Set MyDocument = Application.ActiveWindow
    
    #If Mac Then
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    oPres.Slides.Item(ComboBox1.ListIndex + 1).Copy
    
    oPres.Close
    Set oPres = Nothing
    
    
    MyDocument.Presentation.Slides.Paste
    
    'clean up
    #If Mac Then
        strPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        strPath = Environ("TEMP") & "\"
    #End If
    
    If InsertSlideLibrarySlide.ComboBox1.Enabled = True Then
        
        For i = 1 To InsertSlideLibrarySlide.ComboBox1.ListCount
            Kill (strPath & "tmp.Slide" & i & ".jpg")
        Next i
        
    End If

    
    InsertSlideLibrarySlide.ComboBox1.Enabled = False
    InsertSlideLibrarySlide.Hide
End Sub

Private Sub UserForm_Activate()
    
    
    If GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "") = "" Then
    InsertSlideLibrarySlide.Hide
    MsgBox "No slide library file set. Please set the file in Instrumenta settings."
    
    SettingsForm.Show
    
    Else
    
   
    'Needed to enable OLE automation for this
    
    Dim oPres       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim strPath     As String
    
    #If Mac Then

    
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
    Set oPres = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    #If Mac Then
        strPath = MacScript("return posix path of (path to temporary items) as string")
    #Else
        strPath = Environ("TEMP") & "\"
    #End If
    
    NumberOfSlides = oPres.Slides.Count
    InsertSlideLibrarySlide.ComboBox1.Clear
    
    For Each PresentationSlide In oPres.Slides
        
        PresentationSlide.Export strPath & "tmp.Slide" & PresentationSlide.SlideNumber & ".jpg", "JPG"
        
        SlideTitle = ""
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                SlideTitle = SlidePlaceHolder.TextFrame.TextRange.Text
                Exit For
            End If
        Next SlidePlaceHolder
        
        If SlideTitle = "" Then SlideTitle = PresentationSlide.Name
        InsertSlideLibrarySlide.ComboBox1.AddItem SlideTitle
        
    Next
    
    oPres.Close
    Set oPres = Nothing
    
    InsertSlideLibrarySlide.PreviewImage.Picture = LoadPicture(strPath & "tmp.Slide1.jpg")
    InsertSlideLibrarySlide.ComboBox1.Enabled = True
    InsertSlideLibrarySlide.ComboBox1.ListIndex = 0
    InsertSlideLibrarySlide.Repaint
    
    End If
    
    
End Sub

