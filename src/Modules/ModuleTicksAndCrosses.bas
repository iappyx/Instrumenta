Attribute VB_Name = "ModuleTicksAndCrosses"
Sub TextBulletsTicks()
    
    With Windows(1).Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 252
        .Visible = True
        .Font.Name = "Wingdings"
        .Font.Color = RGB(0, 128, 0)
        
    End With
    
End Sub

Sub TextBulletsCrosses()
    
    With Windows(1).Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 215
        .Visible = True
        .Font.Name = "Calibri"
        .Font.Color = RGB(255, 0, 0)
        
    End With
    
End Sub
