VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetPositionForm 
   Caption         =   "Set position"
   ClientHeight    =   1253
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   3647
   OleObjectBlob   =   "SetPositionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetPositionForm"
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


Private Sub TextBoxLeft_AfterUpdate()

Dim PositionLeft As String
PositionLeft = Replace(CStr(SetPositionForm.TextBoxLeft), ".", GetDecimalSeperator())
PositionLeft = Replace(PositionLeft, ",", GetDecimalSeperator())

If Not PositionLeft = "" Then

If Not Abs(PositionLeft / GetRulerUnitConversion()) > 169000 Then

If Not Application.ActiveWindow.Selection.Type = ppSelectionNone And Not SetPositionForm.TextBoxLeft = "" Then
Application.ActiveWindow.Selection.ShapeRange.left = PositionLeft / GetRulerUnitConversion()
End If

Else

MsgBox "This position is out of bounds"

End If
End If


End Sub


Private Sub TextBoxTop_AfterUpdate()

Dim PositionTop As String
PositionTop = Replace(CStr(SetPositionForm.TextBoxTop), ".", GetDecimalSeperator())
PositionTop = Replace(PositionTop, ",", GetDecimalSeperator())

If Not PositionTop = "" Then

If Not Abs(PositionTop / GetRulerUnitConversion()) > 169000 Then

If Not Application.ActiveWindow.Selection.Type = ppSelectionNone And Not SetPositionForm.TextBoxTop = "" Then
Application.ActiveWindow.Selection.ShapeRange.Top = PositionTop / GetRulerUnitConversion()
End If

Else

MsgBox "This position is out of bounds"

End If

End If

End Sub


Private Sub TextBoxTop_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TextBoxLeft_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 46
    ElseIf KeyAscii = 44 Then
        KeyAscii = 44
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub UserForm_Activate()

RulerTextLeft.Caption = GetRulerUnit()
RulerTextTop.Caption = GetRulerUnit()

Set Sel = Application.ActiveWindow.Selection
    If Sel.Type = ppSelectionShapes Then

        SetPositionForm.TextBoxLeft.Enabled = True
        SetPositionForm.TextBoxTop.Enabled = True
        
        If Sel.ShapeRange.Count > 1 Then
            
            For i = 1 To Sel.ShapeRange.Count
                TotalTop = TotalTop + Sel.ShapeRange(i).Top
                TotalLeft = TotalLeft + Sel.ShapeRange(i).left
            Next i
            
            If Sel.ShapeRange(1).left = TotalLeft / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If Sel.ShapeRange(1).Top = TotalTop / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        End If
        
    ElseIf Sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
        SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        
    Else
        
        SetPositionForm.TextBoxLeft = ""
        SetPositionForm.TextBoxTop = ""
        SetPositionForm.TextBoxLeft.Enabled = False
        SetPositionForm.TextBoxTop.Enabled = False
       
    End If
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


Set Sel = Application.ActiveWindow.Selection
       
    If Sel.Type = ppSelectionShapes Then

        SetPositionForm.TextBoxLeft.Enabled = True
        SetPositionForm.TextBoxTop.Enabled = True
        
        If Sel.ShapeRange.Count > 1 Then
            
            For i = 1 To Sel.ShapeRange.Count
                TotalTop = TotalTop + Sel.ShapeRange(i).Top
                TotalLeft = TotalLeft + Sel.ShapeRange(i).left
            Next i
            
            If Sel.ShapeRange(1).left = TotalLeft / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If Sel.ShapeRange(1).Top = TotalTop / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        End If
        
    ElseIf Sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * GetRulerUnitConversion(), 2)
        SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        
    Else
        
        SetPositionForm.TextBoxLeft = ""
        SetPositionForm.TextBoxTop = ""
        SetPositionForm.TextBoxLeft.Enabled = False
        SetPositionForm.TextBoxTop.Enabled = False
        
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnloadSetPositionAppEventHandler
End Sub
