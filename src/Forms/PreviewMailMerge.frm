Attribute VB_Name = "PreviewMailMerge"
Attribute VB_Base = "0{58F85C29-C944-4484-9547-8BBD65FBCD4A}{96EDEF7F-8BBB-449D-B63F-E9036273A841}"
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

Private Sub CommandButton1_Click()
PreviewMailMerge.Hide
End Sub

Private Sub CommandButton2_Click()
CancelTriggered = True
PreviewMailMerge.Hide
End Sub
