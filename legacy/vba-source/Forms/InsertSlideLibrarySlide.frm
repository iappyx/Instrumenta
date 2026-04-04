Attribute VB_Name = "InsertSlideLibrarySlide"
Attribute VB_Base = "0{921588DB-461A-4D69-A578-17A03108E8B7}{527067C2-41FA-46FA-B518-F83563C7AC88}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

Dim ButtonHandlers  As Collection

Public Sub ToggleCheckBox(ByVal tagValue As Integer)
    Dim ctrl        As control
    For Each Page In MultiPageThumbnailGrid.Pages
        
        For Each ctrl In Page.Controls
            If TypeName(ctrl) = "CheckBox" And ctrl.Tag = CStr(tagValue) Then
                ctrl.value = Not ctrl.value
                Exit For
            End If
        Next ctrl
        
    Next Page
    
    SelectedCount = ReturnSelectedCount
    
    If SelectedCount > 0 Then
        InsertSlideKeepSourceButton.enabled = True
        InsertSlideButton.enabled = True
        
        If SelectedCount = 1 Then
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
            InsertSlideButton.caption = "Insert selected slide"
        Else
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting (" & SelectedCount & ")"
            InsertSlideButton.caption = "Insert selected slides (" & SelectedCount & ")"
        End If
        
    Else
        InsertSlideKeepSourceButton.enabled = False
        InsertSlideButton.enabled = False
        InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
        InsertSlideButton.caption = "Insert selected slide"
    End If
    
End Sub

Private Sub ZoomInButton_Click()
    NumberOfColumns = GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryMaxColumns", 3) - 1
    If NumberOfColumns = 0 Then NumberOfColumns = 1
    SaveSetting "Instrumenta", "SlideLibrary", "SlideLibraryMaxColumns", NumberOfColumns
    RepaintThumbnails (NumberOfColumns)
End Sub

Private Sub ZoomOutButton_Click()
    NumberOfColumns = GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryMaxColumns", 3) + 1
    If NumberOfColumns = 11 Then NumberOfColumns = 10
    SaveSetting "Instrumenta", "SlideLibrary", "SlideLibraryMaxColumns", NumberOfColumns
    RepaintThumbnails (NumberOfColumns)
End Sub

Private Sub InsertSlideKeepSourceButton_Click()
    Dim LibraryPresentation       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim TempPath     As String
    Set MyDocument = Application.ActiveWindow
    Set Testdocument = Application.ActiveWindow.Presentation
    
    #If Mac Then
        Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
        Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    For Each Page In MultiPageThumbnailGrid.Pages
        
        For Each ctrl In Page.Controls
            
            If TypeName(ctrl) = "CheckBox" Then
                If ctrl.value = True Then
                    LibraryPresentation.Slides.Item(CInt(ctrl.Tag)).Copy
                    Testdocument.Windows(1).Activate
                    Testdocument.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")
                    DoEvents
                End If
            End If
        Next ctrl
        
    Next Page
    
    LibraryPresentation.Close
    
    Set LibraryPresentation = Nothing
    
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Private Sub CancelButton_Click()
    
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Private Sub OpenSlideLibrary_Click()
    
    InsertSlideLibrarySlide.Hide
    Unload Me
    OpenSlideLibraryFile
    
End Sub

Private Function ReturnSelectedCount() As Integer
    
    SelectedCount = 0
    
    For Each Page In MultiPageThumbnailGrid.Pages
        
        For Each ctrl In Page.Controls
            
            If TypeName(ctrl) = "CheckBox" Then
                If (ctrl.value = True) Then
                    SelectedCount = SelectedCount + 1
                End If
            End If
        Next ctrl
        
    Next Page
    
    ReturnSelectedCount = SelectedCount
    
End Function

Private Sub SelectAllButton_Click()
    
    For Each ctrl In MultiPageThumbnailGrid.Pages(MultiPageThumbnailGrid.value).Controls
        
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = True
        End If
    Next ctrl
    
    SelectedCount = ReturnSelectedCount
    
    If SelectedCount > 0 Then
        InsertSlideKeepSourceButton.enabled = True
        InsertSlideButton.enabled = True
        
        If SelectedCount = 1 Then
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
            InsertSlideButton.caption = "Insert selected slide"
        Else
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting (" & SelectedCount & ")"
            InsertSlideButton.caption = "Insert selected slides (" & SelectedCount & ")"
        End If
        
    Else
        InsertSlideKeepSourceButton.enabled = False
        InsertSlideButton.enabled = False
        InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
        InsertSlideButton.caption = "Insert selected slide"
    End If
    
End Sub

Private Sub SelectNoneButton_Click()
    
    For Each ctrl In MultiPageThumbnailGrid.Pages(MultiPageThumbnailGrid.value).Controls
        
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = False
        End If
    Next ctrl
    
    SelectedCount = ReturnSelectedCount
    
    If SelectedCount > 0 Then
        InsertSlideKeepSourceButton.enabled = True
        InsertSlideButton.enabled = True
        
        If SelectedCount = 1 Then
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
            InsertSlideButton.caption = "Insert selected slide"
        Else
            InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting (" & SelectedCount & ")"
            InsertSlideButton.caption = "Insert selected slides (" & SelectedCount & ")"
        End If
        
    Else
        InsertSlideKeepSourceButton.enabled = False
        InsertSlideButton.enabled = False
        InsertSlideKeepSourceButton.caption = "Insert selected and maintain source formatting"
        InsertSlideButton.caption = "Insert selected slide"
    End If
    
    #If Mac Then
        InsertSlideKeepSourceButton.enabled = True
        InsertSlideButton.enabled = True
    #End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim TempPath     As String
    
    #If Mac Then
        TempPath = MacScript("return posix path of (path To temporary items) As string")
    #Else
        TempPath = Environ("TEMP") & "\"
    #End If
    
    For Each Page In MultiPageThumbnailGrid.Pages
        
        For Each ctrl In Page.Controls
            
            If TypeName(ctrl) = "CheckBox" Then
                Kill (TempPath & "tmp.Slide" & ctrl.Tag & ".jpg")
            End If
        Next ctrl
        
    Next Page
    
    InsertSlideLibrarySlide.Hide
    Unload Me
End Sub

Private Sub InsertSlideButton_Click()
    Dim LibraryPresentation       As PowerPoint.Presentation
    Dim PresentationSlide As PowerPoint.Slide
    Dim TempPath     As String
    Set MyDocument = Application.ActiveWindow
    
    #If Mac Then
        Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
    #Else
        Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
    #End If
    
    For Each Page In MultiPageThumbnailGrid.Pages
        
        For Each ctrl In Page.Controls
            
            If TypeName(ctrl) = "CheckBox" Then
                If ctrl.value = True Then
                    LibraryPresentation.Slides.Item(CInt(ctrl.Tag)).Copy
                    MyDocument.Presentation.Slides.Paste
                End If
            End If
        Next ctrl
        
    Next Page
    
    LibraryPresentation.Close
    
    Set LibraryPresentation = Nothing
    
    InsertSlideLibrarySlide.Hide
    Unload Me
    
End Sub

Sub RepaintThumbnails(NewThumbnailGridMaxCols As Integer)
    Dim Page        As MSForms.Page
    Dim ctrl        As control
    Dim row         As Integer
    Dim col         As Integer
    Dim ThumbnailWidth As Integer
    Dim ThumbnailHeight As Integer
    Dim i           As Integer
    
    ThumbnailWidth = (650 - (NewThumbnailGridMaxCols * 10)) / NewThumbnailGridMaxCols
    ThumbnailHeight = ThumbnailWidth / 16 * 9
    
    For Each Page In MultiPageThumbnailGrid.Pages
        row = 0
        col = 0
        i = 0
        
        For Each ctrl In Page.Controls
            If TypeName(ctrl) = "Image" Then
                
                ctrl.left = 10 + col * (ThumbnailWidth + 10)
                ctrl.Top = 10 + row * (ThumbnailHeight + 10)
                ctrl.width = ThumbnailWidth
                ctrl.height = ThumbnailHeight
                
                Dim checkCtrl As control
                For Each checkCtrl In Page.Controls
                    If TypeName(checkCtrl) = "CheckBox" And checkCtrl.Tag = ctrl.Tag Then
                        checkCtrl.left = ctrl.left + ctrl.width - 15
                        checkCtrl.Top = ctrl.Top + ctrl.height - 15
                    End If
                Next checkCtrl
                
                Dim buttonCtrl As control
                For Each buttonCtrl In Page.Controls
                    If TypeName(buttonCtrl) = "CommandButton" And buttonCtrl.Tag = ctrl.Tag Then
                        buttonCtrl.left = ctrl.left
                        buttonCtrl.Top = ctrl.Top
                        buttonCtrl.width = ctrl.width
                        buttonCtrl.height = ctrl.height
                    End If
                Next buttonCtrl
                
                col = col + 1
                If col >= NewThumbnailGridMaxCols Then
                    col = 0
                    row = row + 1
                End If
                
                i = i + 1
            End If
        Next ctrl
        
        Page.ScrollHeight = 10 + (row + 1) * (ThumbnailHeight + 10)
        
        Page.Repaint
    Next Page
End Sub

Private Sub UserForm_Activate()
    
    MultiPageThumbnailGrid.visible = False
    If GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "") = "" Then
        InsertSlideLibrarySlide.Hide
        MsgBox "No slide library file set. Please set the file in Infront Toolkit settings."
        SettingsForm.Show
    Else
        
        'Needed to enable OLE automation for this
        Dim LibraryPresentation As PowerPoint.Presentation
        Dim PresentationSlide As PowerPoint.Slide
        Dim TempPath    As String
        
        #If Mac Then
            Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
        #Else
            Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""), , , msoFalse)
        #End If
        
        #If Mac Then
            TempPath = MacScript("return posix path of (path To temporary items) As string")
        #Else
            TempPath = Environ("TEMP") & "\"
        #End If
        
        NumberOfSlides = LibraryPresentation.Slides.count
        slideHeight = 500
        slideWidth = (LibraryPresentation.PageSetup.slideWidth / LibraryPresentation.PageSetup.slideHeight) * slideHeight
        
        Dim Thumbnail         As MSForms.image
        Dim ThumbnailGridMaxCols     As Integer
        Dim ThumbnailWidth  As Integer
        Dim ThumbnailHeight As Integer
        
        ThumbnailGridMaxCols = GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryMaxColumns", 3)
        ThumbnailWidth = (650 - (ThumbnailGridMaxCols * 10)) / ThumbnailGridMaxCols
        ThumbnailHeight = ThumbnailWidth / 16 * 9
        
        MultiPageThumbnailGrid.Pages.Clear
        CurrentPageIndex = 0
        Dim CurrentPage As MSForms.Page
        
        For Each PresentationSlide In LibraryPresentation.Slides
            i = i + 1
            
            CurrentSectionIndex = PresentationSlide.sectionIndex
            
            If CurrentSectionIndex <> CurrentPageIndex Then
                
                row = 0
                col = 0
                
                If LibraryPresentation.SectionProperties.count = 0 Then
                    CurrentSectionName = "Default section" & " (" & LibraryPresentation.Slides.count & ")"
                Else
                    CurrentSectionName = LibraryPresentation.SectionProperties.name(PresentationSlide.sectionIndex) & " (" & LibraryPresentation.SectionProperties.SlidesCount(PresentationSlide.sectionIndex) & ")"
                End If
                
                Set CurrentPage = MultiPageThumbnailGrid.Pages.Add("NewPage" & CurrentSectionIndex, CurrentSectionName)
                CurrentPageIndex = CurrentSectionIndex
                CurrentPage.ScrollBars = fmScrollBarsVertical
            End If
            
            PresentationSlide.Export TempPath & "tmp.Slide" & PresentationSlide.SlideNumber & ".jpg", "JPG", slideWidth, slideHeight
            
            Set Thumbnail = CurrentPage.Controls.Add("Forms.Image.1", "Thumbnail" & i)
            
            With Thumbnail
                .left = 10 + col * (ThumbnailWidth + 10)
                .Top = 10 + row * (ThumbnailHeight + 10)
                .width = ThumbnailWidth
                .height = ThumbnailHeight
                .Picture = LoadPicture(TempPath & "tmp.Slide" & i & ".jpg")
                .Tag = i
                .PictureSizeMode = fmPictureSizeModeZoom
            End With
            
            Set ThumbnailCheck = CurrentPage.Controls.Add("Forms.CheckBox.1", "CheckBox" & i)
            
            With ThumbnailCheck
                .left = Thumbnail.left + Thumbnail.width - 15
                .Top = Thumbnail.Top + Thumbnail.height - 15
                .width = 15
                .height = 15
                .Tag = i
                .BackStyle = fmBackStyleTransparent
            End With
            
            #If Mac Then
            
            'Mac does not support transparent overlay buttons
            
            #Else
            
            Dim ClickOverlay As MSForms.CommandButton
            Set ClickOverlay = CurrentPage.Controls.Add("Forms.CommandButton.1", "ClickOverlay" & i)
            
            With ClickOverlay
                .left = Thumbnail.left
                .Top = Thumbnail.Top
                .width = Thumbnail.width
                .height = Thumbnail.height
                .caption = ""
                .Tag = i
                .BackStyle = fmBackStyleTransparent
            End With
            
            If ButtonHandlers Is Nothing Then Set ButtonHandlers = New Collection
            
            Dim ButtonObj As ThumbnailButtonHandler
            Set ButtonObj = New ThumbnailButtonHandler
            Set ButtonObj.ClickOverlay = ClickOverlay
            
            ButtonHandlers.Add ButtonObj
            
            #End If
            
            col = col + 1
            If col >= ThumbnailGridMaxCols Then
                col = 0
                row = row + 1
            End If
            
            NewScrollHeight = 10 + (row + 1) * (ThumbnailHeight + 10)
            
            ' Repaint before setting new ScrollHeight needed on Mac
            CurrentPage.Repaint
            CurrentPage.ScrollHeight = NewScrollHeight
            CurrentPage.Repaint
        Next
        
        LibraryPresentation.Close
        Set LibraryPresentation = Nothing
        
    End If
    MultiPageThumbnailGrid.visible = True
    
    #If Mac Then
        InsertSlideKeepSourceButton.enabled = True
        InsertSlideButton.enabled = True
    #End If
    
End Sub
