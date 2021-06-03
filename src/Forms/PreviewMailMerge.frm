VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreviewMailMerge 
   Caption         =   "Preview mail merge"
   ClientHeight    =   8113
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10836
   OleObjectBlob   =   "PreviewMailMerge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PreviewMailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
PreviewMailMerge.Hide
End Sub

Private Sub CommandButton2_Click()
CancelTriggered = True
PreviewMailMerge.Hide
End Sub
