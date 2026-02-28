Attribute VB_Name = "SetPositionForm"
Attribute VB_Base = "0{6D6A24A6-BDC9-45C4-A638-190858ED708A}{870691B3-AA43-47BC-BF47-1893B30C99BD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'MIT License
'Copyright (c) 2021 - 2026 iappyx

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

RulerTextLeft.caption = GetRulerUnit()
RulerTextTop.caption = GetRulerUnit()

Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionShapes Then

        SetPositionForm.TextBoxLeft.enabled = True
        SetPositionForm.TextBoxTop.enabled = True
        
        If sel.ShapeRange.count > 1 Then
            
            For i = 1 To sel.ShapeRange.count
                TotalTop = TotalTop + sel.ShapeRange(i).Top
                TotalLeft = TotalLeft + sel.ShapeRange(i).left
            Next i
            
            If sel.ShapeRange(1).left = TotalLeft / sel.ShapeRange.count Then
                SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If sel.ShapeRange(1).Top = TotalTop / sel.ShapeRange.count Then
                SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        End If
        
    ElseIf sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
        SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        
    Else
        
        SetPositionForm.TextBoxLeft = ""
        SetPositionForm.TextBoxTop = ""
        SetPositionForm.TextBoxLeft.enabled = False
        SetPositionForm.TextBoxTop.enabled = False
       
    End If
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)


Set sel = Application.ActiveWindow.Selection
       
    If sel.Type = ppSelectionShapes Then

        SetPositionForm.TextBoxLeft.enabled = True
        SetPositionForm.TextBoxTop.enabled = True
        
        If sel.ShapeRange.count > 1 Then
            
            For i = 1 To sel.ShapeRange.count
                TotalTop = TotalTop + sel.ShapeRange(i).Top
                TotalLeft = TotalLeft + sel.ShapeRange(i).left
            Next i
            
            If sel.ShapeRange(1).left = TotalLeft / sel.ShapeRange.count Then
                SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If sel.ShapeRange(1).Top = TotalTop / sel.ShapeRange.count Then
                SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
            SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        End If
        
    ElseIf sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(sel.ShapeRange.left * GetRulerUnitConversion(), 2)
        SetPositionForm.TextBoxTop = Round(sel.ShapeRange.Top * GetRulerUnitConversion(), 2)
        
    Else
        
        SetPositionForm.TextBoxLeft = ""
        SetPositionForm.TextBoxTop = ""
        SetPositionForm.TextBoxLeft.enabled = False
        SetPositionForm.TextBoxTop.enabled = False
        
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnloadSetPositionAppEventHandler
End Sub
