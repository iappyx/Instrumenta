VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SlidesScopeForm 
   Caption         =   "What is the scope?"
   ClientHeight    =   1110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "SlidesScopeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SlidesScopeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
