Attribute VB_Name = "AboutDialog"
Attribute VB_Base = "0{7544FB23-B219-4C47-A414-67D23D550214}{B3C027F4-8A73-49BA-811F-B81D5812FD5C}"
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

Private Sub Label2_Click()
    Dim URL As String
    Dim tempPresentation As Presentation
        
    URL = "https://github.com/iappyx/Instrumenta"
    
    If Presentations.count = 0 Then
        Set tempPresentation = Presentations.Add
        tempPresentation.FollowHyperlink URL
        tempPresentation.Close
    Else
        ActivePresentation.FollowHyperlink URL
    End If
    
    
End Sub

Private Sub Label3_Click()

    Dim URL As String
    Dim tempPresentation As Presentation


    URL = "https://github.com/iappyx/Instrumenta/blob/main/v/" & InstrumentaVersion & ".md"

    If Presentations.count = 0 Then
        Set tempPresentation = Presentations.Add
        tempPresentation.FollowHyperlink URL
        tempPresentation.Close
    Else
        ActivePresentation.FollowHyperlink URL
    End If


End Sub
