VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageTagsForm 
   Caption         =   "Manage tags"
   ClientHeight    =   5130
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   15043
   OleObjectBlob   =   "ManageTagsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageTagsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub CommandButton1_Click()
    ManageTagsForm.Hide
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    AddTag
End Sub

Private Sub CommandButton3_Click()
    DeleteTag
End Sub

Private Sub CommandButton4_Click()
    DeleteAllTags
End Sub

Private Sub CommandButton5_Click()
AddSpecialSlideTag "filename"
ManageTagsForm.Hide
ShowFormManageTags
End Sub

Private Sub CommandButton6_Click()
AddSpecialSlideTag "slidenum"
ManageTagsForm.Hide
ShowFormManageTags
End Sub
