VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JsonInputForm 
   Caption         =   "Import JSON"
   ClientHeight    =   9090.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
   OleObjectBlob   =   "JsonInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JsonInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Result As String
Public Cancelled As Boolean

Private Sub cmdOK_Click()
    Result = Me.txtJson.text
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

