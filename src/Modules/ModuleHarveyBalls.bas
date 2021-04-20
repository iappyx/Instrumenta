Attribute VB_Name = "ModuleHarveyBalls"
Sub GenerateHarveyBallPercent(FillPercentage As Double)
    #If Mac Then
        MsgBox "This Function will Not yet work On a Mac"
    #Else
        Dim HarveyCircle, HarveyFill As Shape
        
        Set HarveyCircle = ActivePresentation.Slides(1).Shapes.AddShape(msoShapeOval, 100, 100, 50, 50)
        Set HarveyFill = ActivePresentation.Slides(1).Shapes.AddShape(msoShapePie, 101, 101, 48, 48)
        With HarveyFill
            .Adjustments.Item(2) = -90
            .Adjustments.Item(1) = ((FillPercentage / 100) * 360) - 90
            .Line.Visible = False
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        With HarveyCircle
            .Line.Visible = False
        End With
        
        If FillPercentage > 0 Then
            HarveyFill.Adjustments(1) = HarveyFill.Adjustments(1) - 0.1
        End If
        
        'Code below does not work on mac - needs fix
        Call ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array(HarveyCircle.ZOrderPosition, HarveyFill.ZOrderPosition)).MergeShapes(msoMergeCombine)
    #End If
End Sub
