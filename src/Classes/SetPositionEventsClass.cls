VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetPositionEventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1


Private Sub App_AfterShapeSizeChange(ByVal shp As shape)
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

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    
    
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

