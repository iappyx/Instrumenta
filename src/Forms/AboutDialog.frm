VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutDialog 
   Caption         =   "About"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270.001
   OleObjectBlob   =   "AboutDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
    Dim URL As String
    URL = "https://github.com/iappyx/Instrumenta"
    ActivePresentation.FollowHyperlink URL
End Sub
