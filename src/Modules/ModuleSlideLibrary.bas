Attribute VB_Name = "ModuleSlideLibrary"
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

Sub OpenSlideLibraryFile()

If GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "") = "" Then
    
    MsgBox "No slide library file set. Please set the file in Instrumenta settings."
    
    SettingsForm.Show
    
Else
    Dim LibraryPresentation       As PowerPoint.Presentation
    Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
End If


End Sub


Sub AddSelectedSlidesToLibraryFile()

Dim LibraryPresentation       As PowerPoint.Presentation
Set MyDocument = Application.ActiveWindow

If GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "") = "" Then
    
    MsgBox "No slide library file set. Please set the file in Instrumenta settings."
    
    SettingsForm.Show
    
Else
    
    If Application.ActiveWindow.Selection.SlideRange.Count > 0 Then
        
        Set LibraryPresentation = PowerPoint.Presentations.Open(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))
        
        MyDocument.Selection.SlideRange.Copy
        
        LibraryPresentation.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")
                
        LibraryPresentation.SaveCopyAs GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "")
        LibraryPresentation.Close
        
        Set LibraryPresentation = Nothing
        
    Else
        
        MsgBox "No slides selected."
        
    End If
    
End If


End Sub
