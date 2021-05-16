VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectSlidesByTagForm 
   Caption         =   "Select slides by tag or stamp"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8175
   OleObjectBlob   =   "SelectSlidesByTagForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectSlidesByTagForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
SelectSlidesByTag
End Sub

Private Sub CommandButton2_Click()
SelectSlidesByStamp StampComboBox.Value
End Sub

Private Sub CommandButton3_Click()
SelectSlidesByTagForm.Hide
End Sub

Private Sub SlideTagComboBox_Change()
PopulateSlideTagValueListbox
End Sub
