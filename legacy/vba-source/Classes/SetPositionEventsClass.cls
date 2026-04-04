Attribute VB_Name = "SetPositionEventsClass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1


Private Sub App_AfterShapeSizeChange(ByVal shp As shape)
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

Private Sub app_WindowSelectionChange(ByVal sel As Selection)
    
    
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

