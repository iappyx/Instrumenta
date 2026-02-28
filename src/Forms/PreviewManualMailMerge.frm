Attribute VB_Name = "PreviewManualMailMerge"
Attribute VB_Base = "0{0E056369-684B-4D07-98FF-E72076BFC1BF}{E516F3D0-E591-40CC-868F-7EE8585C1501}"
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
PreviewManualMailMerge.Hide
End Sub

Private Sub CommandButton2_Click()
CancelTriggered = True
PreviewManualMailMerge.Hide
End Sub

Private Sub CommandButton3_Click()
PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 1) = ReplaceTextTextBox.text
End Sub

Private Sub MailMergeListBox_Click()

If Not PreviewManualMailMerge.MailMergeListBox.ListIndex = -1 Then
ReplaceTextTextBox.text = PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 1)
ReplaceTextFrame.caption = "Replace " & PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 0) & " with:"
End If


End Sub
