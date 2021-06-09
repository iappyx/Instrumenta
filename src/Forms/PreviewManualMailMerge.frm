VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreviewManualMailMerge 
   Caption         =   "Preview manual mail merge"
   ClientHeight    =   8115
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   10890
   OleObjectBlob   =   "PreviewManualMailMerge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PreviewManualMailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
PreviewManualMailMerge.Hide
End Sub

Private Sub CommandButton2_Click()
CancelTriggered = True
PreviewManualMailMerge.Hide
End Sub

Private Sub CommandButton3_Click()
PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 1) = ReplaceTextTextBox.Text
End Sub

Private Sub MailMergeListBox_Click()

If Not PreviewManualMailMerge.MailMergeListBox.ListIndex = -1 Then
ReplaceTextTextBox.Text = PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 1)
ReplaceTextFrame.Caption = "Replace " & PreviewManualMailMerge.MailMergeListBox.List(PreviewManualMailMerge.MailMergeListBox.ListIndex, 0) & " with:"
End If


End Sub
