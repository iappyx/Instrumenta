Attribute VB_Name = "ModuleFunctions"
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

Function RemoveDuplicates(InputArray) As Variant
    
    Dim OutputArray, InputValue, OutputValue As Variant
    Dim MatchFound  As Boolean
    
    On Error Resume Next
    OutputArray = Array("")
    For Each InputValue In InputArray
        MatchFound = False
        
        If IsEmpty(InputValue) Then GoTo ForceNext
        For Each OutputValue In OutputArray
            If OutputValue = InputValue Then
                MatchFound = True
                Exit For
            End If
        Next OutputValue
        
        If MatchFound = False Then
            ReDim Preserve OutputArray(UBound(OutputArray, 1) + 1)
            OutputArray(UBound(OutputArray, 1) - 1) = InputValue
        End If
        
ForceNext:
    Next
    RemoveDuplicates = OutputArray
    
End Function

Sub SetProgress(PercentageCompleted As Single)

    ProgressForm.ProgressBar.Width = PercentageCompleted * 2
    ProgressForm.ProgressLabel.Caption = Round(PercentageCompleted, 0) & "% completed"
    DoEvents
    
End Sub

Function MacFileDialog(filepath As String) As String
  MacFileDialogMacScript = "set applescript's text item delimiters to "","" " & vbNewLine & "try " & vbNewLine & "set selectedFile to (choose file " & _
    "with prompt ""Please select a file"" default location alias """ & filepath & """ multiple selections allowed false) as string" & vbNewLine & "set applescript's text item delimiters to """" " & vbNewLine & _
    "on error errStr number errorNumber" & vbNewLine & "return errorNumber " & vbNewLine & "end try " & vbNewLine & "return selectedFile"
  MacFileDialog = MacScript(MacFileDialogMacScript)
  
  If MacFileDialog = "-128" Then
  MacFileDialog = ""
  Else
      If CInt(Split(Application.Version, ".")(0)) >= 15 Then
    MacFileDialog = Replace(MacFileDialog, ":", "/")
    MacFileDialog = Replace(MacFileDialog, "Macintosh HD", "", Count:=1)
        End If
  End If
  
End Function

'This does not work well in all cases
'Function MacSendMailViaOutlook(subject As String, filepath As String)
'MacSendMailViaOutlookMacScript = "tell application ""Microsoft Outlook""" & vbNewLine & "set NewMail to (make new outgoing message with properties {subject:""" & subject & """})" & vbNewLine & _
'"tell NewMail" & vbNewLine & "set AttachmentPath to POSIX file """ & filepath & """" & vbNewLine & "make new attachment with properties {file:AttachmentPath as alias}" & vbNewLine & "Delay 0.5" & vbNewLine & _
'"end tell" & vbNewLine & "open NewMail" & vbNewLine & "Activate NewMail" & vbNewLine & "end tell" & vbNewLine & "return ""Done"""
'MacSendMailViaOutlook = MacScript(MacSendMailViaOutlookMacScript)
'End Function

