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

Private Sub TextBoxLeft_AfterUpdate()

If Not Application.ActiveWindow.Selection.Type = ppSelectionNone And Not SetPositionForm.TextBoxLeft = "" Then
Application.ActiveWindow.Selection.ShapeRange.left = SetPositionForm.TextBoxLeft / 0.0352777778
End If

End Sub

Private Sub TextBoxTop_AfterUpdate()

If Not Application.ActiveWindow.Selection.Type = ppSelectionNone And Not SetPositionForm.TextBoxTop = "" Then
Application.ActiveWindow.Selection.ShapeRange.Top = SetPositionForm.TextBoxTop / 0.0352777778
End If

End Sub

Private Sub UserForm_Activate()
    
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
                SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If Sel.ShapeRange(1).Top = TotalTop / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
            SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
        End If
        
    ElseIf Sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
        SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
        
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
                SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
            Else
                SetPositionForm.TextBoxLeft = ""
            End If
            
            If Sel.ShapeRange(1).Top = TotalTop / Sel.ShapeRange.Count Then
                SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
            Else
                SetPositionForm.TextBoxTop = ""
            End If
            
        Else
            SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
            SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
        End If
        
    ElseIf Sel.Type = ppSelectionText Then
        
        SetPositionForm.TextBoxLeft = Round(Sel.ShapeRange.left * 0.0352777778, 2)
        SetPositionForm.TextBoxTop = Round(Sel.ShapeRange.Top * 0.0352777778, 2)
        
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
