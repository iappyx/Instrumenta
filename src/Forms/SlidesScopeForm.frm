Attribute VB_Name = "SlidesScopeForm"
Attribute VB_Base = "0{3A332116-CA50-4087-9C60-95F734E6B53C}{4B477874-2F38-4979-8B3B-5E47E5D44E57}"
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

Public UserChoice As String

Private Sub AllSlidesButton_Click()
    UserChoice = "all"
    Me.Hide
    
End Sub

Private Sub CancelButton_Click()
    UserChoice = "cancel"
    Me.Hide
    
End Sub

Private Sub SelectedSlidesButton_Click()
    UserChoice = "selected"
    Me.Hide
End Sub
