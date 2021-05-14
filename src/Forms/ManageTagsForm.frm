VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageTagsForm 
   Caption         =   "Manage tags"
   ClientHeight    =   6645
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   9870.001
   OleObjectBlob   =   "ManageTagsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageTagsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ManageTagsForm.Hide
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
