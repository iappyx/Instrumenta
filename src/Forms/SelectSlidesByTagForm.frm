Attribute VB_Name = "SelectSlidesByTagForm"
Attribute VB_Base = "0{CC3DF6EE-2205-455D-B5AA-81223D7E0CD6}{246F57BB-2D0B-4FF8-AA89-4B0D22C259DC}"
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
SelectSlidesByTag
End Sub

Private Sub CommandButton2_Click()
SelectSlidesByStamp StampComboBox.value
End Sub

Private Sub CommandButton3_Click()
SelectSlidesByTagForm.Hide
Unload Me
End Sub

Private Sub SlideTagComboBox_Change()
PopulateSlideTagValueListbox
End Sub
