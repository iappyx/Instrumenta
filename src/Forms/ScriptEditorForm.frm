VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScriptEditorForm 
   Caption         =   "Instrumenta script editor"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16560
   OleObjectBlob   =   "ScriptEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScriptEditorForm"
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


Option Explicit

Private Sub btnRun_Click()
    If Trim(txtScript.text) = "" Then
        MsgBox "Please enter a script first.", vbInformation
        Exit Sub
    End If
    
    txtLog.text = ""
    
    RunScript txtScript.text
    
    Dim msg As Variant
    Dim logText As String
    For Each msg In ScriptLog
        logText = logText & msg & vbCrLf
    Next msg
    txtLog.text = logText
    
    txtLog.SelStart = 0
End Sub

Private Sub btnClear_Click()
    If MsgBox("Clear the script?", vbQuestion + vbYesNo) = vbYes Then
        txtScript.text = ""
        txtScript.SetFocus
    End If
End Sub

Private Sub btnClearLog_Click()
    txtLog.text = ""
End Sub

Private Sub btnExample_Click()
    Dim example As String
    example = "# Instrumenta Script Example" & vbCrLf
    example = example & "# Lines starting with # are comments" & vbCrLf
    example = example & "" & vbCrLf
    example = example & "# Make script re-runnable by cleaning up first" & vbCrLf
    example = example & "DELETE WHERE name STARTSWITH ""script_""" & vbCrLf
    example = example & "" & vbCrLf
    example = example & "# Insert a rectangle and style it" & vbCrLf
    example = example & "INSERT RECTANGLE AT 50, 50 WIDTH 300 HEIGHT 200 NAME ""script_box""" & vbCrLf
    example = example & "SET fill.color = #003366" & vbCrLf
    example = example & "SET font.color = #FFFFFF" & vbCrLf
    example = example & "" & vbCrLf
    example = example & "# Insert a title textbox" & vbCrLf
    example = example & "INSERT TEXTBOX AT 60, 60 WIDTH 280 HEIGHT 40 NAME ""script_title"" TEXT ""My Title""" & vbCrLf
    example = example & "SET font.size = 18" & vbCrLf
    example = example & "SET font.bold = TRUE" & vbCrLf
    example = example & "" & vbCrLf
    example = example & "# Select shapes by name prefix and call an Instrumenta function" & vbCrLf
    example = example & "SELECT WHERE name STARTSWITH ""script_""" & vbCrLf
    example = example & "CALL ObjectsAlignTops" & vbCrLf
    example = example & "" & vbCrLf
    example = example & "# After CALL, re-sync working set explicitly if needed" & vbCrLf
    example = example & "USE SELECTION" & vbCrLf
    example = example & "SET font.name = ""Calibri""" & vbCrLf
    
    If Trim(txtScript.text) <> "" Then
        If MsgBox("Replace current script with example?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    txtScript.text = example
    txtScript.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub Label3_Click()
    Dim URL As String
    Dim tempPresentation As Presentation


    URL = "https://github.com/iappyx/Instrumenta/blob/main/SCRIPT.MD"

    If Presentations.count = 0 Then
        Set tempPresentation = Presentations.Add
        tempPresentation.FollowHyperlink URL
        tempPresentation.Close
    Else
        ActivePresentation.FollowHyperlink URL
    End If
End Sub

Private Sub UserForm_Initialize()
    
    txtScript.text = "# Type your script here" & vbCrLf & "# Example: SELECT ALL"
    txtLog.text = ""
    txtLog.Locked = True
    
        Dim codeFontName As String
    #If Mac Then
        codeFontName = "Courier New"
    #Else
        codeFontName = "Consolas"
    #End If
    
    txtScript.Font.name = codeFontName
    txtScript.Font.Size = 8
    txtLog.Font.name = codeFontName
    txtLog.Font.Size = 8
    
End Sub
