VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Settings"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
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

Private Sub UserForm_Activate()

    ShapeStepSizeMargin = GetSetting("Instrumenta", "Shapes", "ShapeStepSizeMargin", "0.2")
    TableStepSizeMargin = GetSetting("Instrumenta", "Tables", "TableStepSizeMargin", "0.2")
    TableStepSizeColumnGaps = GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1.0")
    TableStepSizeRowGaps = GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1.0")
    StickyNotesDefaultText = GetSetting("Instrumenta", "StickyNotes", "StickyNotesDefaultText", "Note")
    
    If GetSetting("Instrumenta", "General", "OperatingMode", "pro") = "pro" Then
    OptionButton1.Value = True
    Else
    OptionButton2.Value = True
    End If

End Sub

Private Sub CancelButton_Click()
    SettingsForm.Hide
End Sub

Private Sub ClearSettingsButton_Click()
    SettingsForm.Hide
    DeleteAllInstrumentaSettings
    SettingsForm.Show
End Sub

Private Sub SaveSettingsButton_Click()

    If ShapeStepSizeMargin = "" Or ShapeStepSizeMargin Like "*[!0-9.]*" Then
        MsgBox ("Please enter data in the following format ###.#")
        ShapeStepSizeMargin.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeMargin = "" Or TableStepSizeMargin Like "*[!0-9.]*" Then
        MsgBox ("Please enter data in the following format ###.#")
        TableStepSizeMargin.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeColumnGaps = "" Or TableStepSizeColumnGaps Like "*[!0-9.]*" Then
        MsgBox ("Please enter data in the following format ###.#")
        TableStepSizeColumnGaps.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeRowGaps = "" Or TableStepSizeRowGaps Like "*[!0-9.]*" Then
        MsgBox ("Please enter data in the following format ###.#")
        TableStepSizeRowGaps.SetFocus
        Exit Sub
    End If
    
    SaveSetting "Instrumenta", "Shapes", "ShapeStepSizeMargin", ShapeStepSizeMargin
    SaveSetting "Instrumenta", "Tables", "TableStepSizeMargin", TableStepSizeMargin
    SaveSetting "Instrumenta", "Tables", "TableStepSizeColumnGaps", TableStepSizeColumnGaps
    SaveSetting "Instrumenta", "Tables", "TableStepSizeRowGaps", TableStepSizeRowGaps
    SaveSetting "Instrumenta", "StickyNotes", "StickyNotesDefaultText", StickyNotesDefaultText
      
    If OptionButton2.Value = True Then
     SaveSetting "Instrumenta", "General", "OperatingMode", "review"
     Call InstrumentaRefresh(UpdateTag:="*R*")
    Else
     SaveSetting "Instrumenta", "General", "OperatingMode", "pro"
     Call InstrumentaRefresh(UpdateTag:="*")
    End If
    
    SettingsForm.Hide
        
End Sub

Private Sub ShapeStepSizeMargin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TableStepSizeMargin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TableStepSizeColumnGaps_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub


Private Sub TableStepSizeRowGaps_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

