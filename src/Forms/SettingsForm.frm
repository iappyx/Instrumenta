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

    ShapeStepSizeMargin = GetSetting("Instrumenta", "Shapes", "ShapeStepSizeMargin", "0" + GetDecimalSeperator() + "2")
    TableStepSizeMargin = GetSetting("Instrumenta", "Tables", "TableStepSizeMargin", "0" + GetDecimalSeperator() + "2")
    TableStepSizeColumnGaps = GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1" + GetDecimalSeperator() + "0")
    TableStepSizeRowGaps = GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1" + GetDecimalSeperator() + "0")
    
    StickyNotesDefaultText = GetSetting("Instrumenta", "StickyNotes", "StickyNotesDefaultText", "Note")
    StickyNotesColor = GetSetting("Instrumenta", "StickyNotes", "StickyNotesColor", "49407")
    StickyNoteColorButton.BackColor = StickyNotesColor
    
    ConfidentialColor = GetSetting("Instrumenta", "Stamps", "ConfidentialColor", "192")
    ConfidentialColorButton.BackColor = ConfidentialColor
    
    DoNotDistributeColor = GetSetting("Instrumenta", "Stamps", "DoNotDistributeColor", "192")
    DoNotDistributeColorButton.BackColor = DoNotDistributeColor
    
    DraftColor = GetSetting("Instrumenta", "Stamps", "DraftColor", "12611584")
    DraftColorButton.BackColor = DraftColor
    
    newColor = GetSetting("Instrumenta", "Stamps", "NewColor", "5287936")
    NewColorButton.BackColor = newColor
    
    ToAppendixColor = GetSetting("Instrumenta", "Stamps", "ToAppendixColor", "8355711")
    ToAppendixColorButton.BackColor = ToAppendixColor
    
    ToBeRemovedColor = GetSetting("Instrumenta", "Stamps", "ToBeRemovedColor", "179")
    ToBeRemovedColorButton.BackColor = ToBeRemovedColor
    
    UpdatedColor = GetSetting("Instrumenta", "Stamps", "UpdatedColor", "39423")
    UpdatedColorButton.BackColor = UpdatedColor
    
    
    SlideLibraryFile = GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "")
    
    
    
    If GetSetting("Instrumenta", "General", "OperatingMode", "default") = "pro" Then
    OptionButton1.Value = True
    ElseIf GetSetting("Instrumenta", "General", "OperatingMode", "default") = "review" Then
    OptionButton2.Value = True
    ElseIf GetSetting("Instrumenta", "General", "OperatingMode", "default") = "default" Then
    OptionButton3.Value = True
    End If
    
    CheckBox1.Value = CBool(GetSetting("Instrumenta", "General", "ContextualButtons", "False"))
    
    RulerUnitsComboBox.Clear
    RulerUnitsComboBox.AddItem ("Inches")
    RulerUnitsComboBox.AddItem ("Centimeters")
    RulerUnitsComboBox.AddItem ("Milimeters")
    RulerUnitsComboBox.AddItem ("Points")
    RulerUnitsComboBox.ListIndex = GetSetting("Instrumenta", "RulerUnits", "ShapePositioning", "1")
    
    DefaultAlignmentMethodComboBox.Clear
    DefaultAlignmentMethodComboBox.AddItem ("Default (based on position)")
    DefaultAlignmentMethodComboBox.AddItem ("To first selected shape")
    DefaultAlignmentMethodComboBox.AddItem ("To last selected shape")
    DefaultAlignmentMethodComboBox.ListIndex = GetSetting("Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", "0")
    
    DefaultTransformationMethodComboBox.Clear
    DefaultTransformationMethodComboBox.AddItem ("Based on first selected shape")
    DefaultTransformationMethodComboBox.AddItem ("Based on last selected shape")
    DefaultTransformationMethodComboBox.ListIndex = GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0")

End Sub

Private Sub CancelButton_Click()
    SettingsForm.Hide
    Unload Me
End Sub

Private Sub ClearSettingsButton_Click()
    SettingsForm.Hide
    DeleteAllInstrumentaSettings
    SettingsForm.Show
End Sub

Private Sub SaveSettingsButton_Click()

    If ShapeStepSizeMargin = "" Or ShapeStepSizeMargin Like "*[!0-9" + GetDecimalSeperator() + "]*" Then
        MsgBox ("Please enter data in the following format #" + GetDecimalSeperator() + "#")
        ShapeStepSizeMargin.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeMargin = "" Or TableStepSizeMargin Like "*[!0-9" + GetDecimalSeperator() + "]*" Then
        MsgBox ("Please enter data in the following format #" + GetDecimalSeperator() + "#")
        TableStepSizeMargin.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeColumnGaps = "" Or TableStepSizeColumnGaps Like "*[!0-9" + GetDecimalSeperator() + "]*" Then
        MsgBox ("Please enter data in the following format #" + GetDecimalSeperator() + "#")
        TableStepSizeColumnGaps.SetFocus
        Exit Sub
    End If
    
        If TableStepSizeRowGaps = "" Or TableStepSizeRowGaps Like "*[!0-9" + GetDecimalSeperator() + "]*" Then
        MsgBox ("Please enter data in the following format #" + GetDecimalSeperator() + "#")
        TableStepSizeRowGaps.SetFocus
        Exit Sub
    End If
    
    SaveSetting "Instrumenta", "Shapes", "ShapeStepSizeMargin", ShapeStepSizeMargin
    SaveSetting "Instrumenta", "Tables", "TableStepSizeMargin", TableStepSizeMargin
    SaveSetting "Instrumenta", "Tables", "TableStepSizeColumnGaps", TableStepSizeColumnGaps
    SaveSetting "Instrumenta", "Tables", "TableStepSizeRowGaps", TableStepSizeRowGaps
    SaveSetting "Instrumenta", "StickyNotes", "StickyNotesDefaultText", StickyNotesDefaultText
    SaveSetting "Instrumenta", "SlideLibrary", "SlideLibraryFile", SlideLibraryFile
    SaveSetting "Instrumenta", "RulerUnits", "ShapePositioning", RulerUnitsComboBox.ListIndex
    SaveSetting "Instrumenta", "AlignDistributeSize", "DefaultAlignmentMethod", DefaultAlignmentMethodComboBox.ListIndex
    SaveSetting "Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", DefaultTransformationMethodComboBox.ListIndex
    
    
    
    red = StickyNoteColorButton.BackColor Mod 256
    green = StickyNoteColorButton.BackColor \ 256 Mod 256
    blue = StickyNoteColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "StickyNotes", "StickyNotesColor", RGB(red, green, blue)
    
    red = ConfidentialColorButton.BackColor Mod 256
    green = ConfidentialColorButton.BackColor \ 256 Mod 256
    blue = ConfidentialColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "ConfidentialColor", RGB(red, green, blue)
    
    red = DoNotDistributeColorButton.BackColor Mod 256
    green = DoNotDistributeColorButton.BackColor \ 256 Mod 256
    blue = DoNotDistributeColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "DoNotDistributeColor", RGB(red, green, blue)
    
    red = DraftColorButton.BackColor Mod 256
    green = DraftColorButton.BackColor \ 256 Mod 256
    blue = DraftColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "DraftColor", RGB(red, green, blue)
    
    red = NewColorButton.BackColor Mod 256
    green = NewColorButton.BackColor \ 256 Mod 256
    blue = NewColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "NewColor", RGB(red, green, blue)
    
    red = ToAppendixColorButton.BackColor Mod 256
    green = ToAppendixColorButton.BackColor \ 256 Mod 256
    blue = ToAppendixColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "ToAppendixColor", RGB(red, green, blue)
    
    red = ToBeRemovedColorButton.BackColor Mod 256
    green = ToBeRemovedColorButton.BackColor \ 256 Mod 256
    blue = ToBeRemovedColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "ToBeRemovedColor", RGB(red, green, blue)
    
    red = UpdatedColorButton.BackColor Mod 256
    green = UpdatedColorButton.BackColor \ 256 Mod 256
    blue = UpdatedColorButton.BackColor \ 65536 Mod 256
    
    SaveSetting "Instrumenta", "Stamps", "UpdatedColor", RGB(red, green, blue)
    SaveSetting "Instrumenta", "General", "ContextualButtons", CStr(CheckBox1.Value)
    DoEvents
          
    If OptionButton2.Value = True Then
     SaveSetting "Instrumenta", "General", "OperatingMode", "review"
     Call InstrumentaRefresh(UpdateTag:="*R*")
    ElseIf OptionButton1.Value = True Then
     SaveSetting "Instrumenta", "General", "OperatingMode", "pro"
     Call InstrumentaRefresh(UpdateTag:="*")
    ElseIf OptionButton3.Value = True Then
     SaveSetting "Instrumenta", "General", "OperatingMode", "default"
     Call InstrumentaRefresh(UpdateTag:="*")
    End If
    

    
    SettingsForm.Hide
    Unload Me
    
        
End Sub

Private Sub StickyNoteColorButton_Click()

red = StickyNoteColorButton.BackColor Mod 256
green = StickyNoteColorButton.BackColor \ 256 Mod 256
blue = StickyNoteColorButton.BackColor \ 65536 Mod 256

StickyNoteColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub


Private Sub ConfidentialColorButton_Click()

red = ConfidentialColorButton.BackColor Mod 256
green = ConfidentialColorButton.BackColor \ 256 Mod 256
blue = ConfidentialColorButton.BackColor \ 65536 Mod 256

ConfidentialColorButton.BackColor = ColorDialog(RGB(red, green, blue))


End Sub

Private Sub DoNotDistributeColorButton_Click()

red = DoNotDistributeColorButton.BackColor Mod 256
green = DoNotDistributeColorButton.BackColor \ 256 Mod 256
blue = DoNotDistributeColorButton.BackColor \ 65536 Mod 256

DoNotDistributeColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub DraftColorButton_Click()

red = DraftColorButton.BackColor Mod 256
green = DraftColorButton.BackColor \ 256 Mod 256
blue = DraftColorButton.BackColor \ 65536 Mod 256

DraftColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub NewColorButton_Click()

red = NewColorButton.BackColor Mod 256
green = NewColorButton.BackColor \ 256 Mod 256
blue = NewColorButton.BackColor \ 65536 Mod 256

NewColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub ToAppendixColorButton_Click()

red = ToAppendixColorButton.BackColor Mod 256
green = ToAppendixColorButton.BackColor \ 256 Mod 256
blue = ToAppendixColorButton.BackColor \ 65536 Mod 256

ToAppendixColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub ToBeRemovedColorButton_Click()

red = ToBeRemovedColorButton.BackColor Mod 256
green = ToBeRemovedColorButton.BackColor \ 256 Mod 256
blue = ToBeRemovedColorButton.BackColor \ 65536 Mod 256

ToBeRemovedColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub UpdatedColorButton_Click()

red = UpdatedColorButton.BackColor Mod 256
green = UpdatedColorButton.BackColor \ 256 Mod 256
blue = UpdatedColorButton.BackColor \ 65536 Mod 256

UpdatedColorButton.BackColor = ColorDialog(RGB(red, green, blue))

End Sub

Private Sub SelectFileButton_Click()
        #If Mac Then
            
            LibraryFile = MacFileDialog("/")
            
            If LibraryFile = "" Then
                MsgBox "No file selected."
                Exit Sub
            End If
            
        #Else
            With Application.FileDialog(msoFileDialogFilePicker)
                .AllowMultiSelect = False
                .Filters.Add "Powerpoint files", "*.pptx; *.ppt", 1
                .Show
                
                If .SelectedItems.count = 0 Then
                    MsgBox "No file selected."
                    Exit Sub
                Else
                    LibraryFile = .SelectedItems.Item(1)
                End If
                
            End With
        #End If
        
        SlideLibraryFile = LibraryFile
End Sub

Private Sub ClearSlideLibraryButton_Click()
SlideLibraryFile = ""
End Sub

Private Sub ShapeStepSizeMargin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TableStepSizeMargin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TableStepSizeColumnGaps_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub


Private Sub TableStepSizeRowGaps_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

