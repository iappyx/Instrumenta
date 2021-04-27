Attribute VB_Name = "ModuleRAGStatus"
Sub GenerateRAGStatus(RagStatus As String)
    
    Set myDocument = Application.ActiveWindow
    RandomNumber = Round(Rnd() * 1000000, 0)
    
        Set RAGBackground = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 100, 100, 94, 34)
        
        With RAGBackground
            .Line.Visible = False
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Name = "RAGBackground" + Str(RandomNumber)
        End With
        
        
        Set GreenStatus = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 104, 104, 26, 26)
        
        With GreenStatus
            .Line.Visible = False
            
            If RagStatus = "green" Then
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "GreenStatus" + Str(RandomNumber)
        End With
    
        Set AmberStatus = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 134, 104, 26, 26)
        
        With AmberStatus
            .Line.Visible = False

            If RagStatus = "amber" Then
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "AmberStatus" + Str(RandomNumber)
        End With
    
        Set RedStatus = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 164, 104, 26, 26)
        
        With RedStatus
            .Line.Visible = False
            
            If RagStatus = "red" Then
            .Fill.ForeColor.RGB = RGB(192, 0, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "RedStatus" + Str(RandomNumber)
        End With
        
        
        ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("RAGBackground" + Str(RandomNumber), "GreenStatus" + Str(RandomNumber), "AmberStatus" + Str(RandomNumber), "RedStatus" + Str(RandomNumber))).Group
    
End Sub
