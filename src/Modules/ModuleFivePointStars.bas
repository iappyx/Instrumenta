Attribute VB_Name = "ModuleFivePointStars"
Sub GenerateFivePointStars(NumberOfStars As Double)
    
    Set myDocument = Application.ActiveWindow
    Dim StarsCount  As Double
    Dim StarsArray  As Variant
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    For StarsCount = 1 To 5
        Set FivePointStar = myDocument.Selection.SlideRange.Shapes.AddShape(msoShape5pointStar, 100 + (StarsCount * 26), 100, 26, 26)
        
        With FivePointStar
            .Line.Visible = False
            .Fill.ForeColor.RGB = RGB(242, 242, 242)
            .Name = "FivePointStar" + Str(StarsCount) + Str(RandomNumber)
        End With
        
        If StarsCount = 1 Then
            StarsArray = Array("FivePointStar" + Str(StarsCount) + Str(RandomNumber))
        Else
            ReDim Preserve StarsArray(UBound(StarsArray) + 1)
            StarsArray(UBound(StarsArray)) = "FivePointStar" + Str(StarsCount) + Str(RandomNumber)
        End If
        
    Next
    
    For StarsCount = 1 To 5
        
        If StarsCount < NumberOfStars + 1 Then
            
            Set FivePointStar = myDocument.Selection.SlideRange.Shapes.AddShape(msoShape5pointStar, 100 + (StarsCount * 26), 100, 26, 26)
            
            With FivePointStar
                .Line.Visible = False
                .Fill.ForeColor.RGB = RGB(255, 192, 0)
                .Name = "FivePointStar" + Str(StarsCount + 5) + Str(RandomNumber)
            End With
            
            If NumberOfStars < StarsCount Then
                Set HalfOfFivePointStar = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, 113 + (StarsCount * 26), 100, 26, 26)
                
                With HalfOfFivePointStar
                    .Name = "HalfFivePointStar" + Str(StarsCount + 5) + Str(RandomNumber)
                End With
                
                ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("FivePointStar" + Str(StarsCount + 5) + Str(RandomNumber), "HalfFivePointStar" + Str(StarsCount + 5) + Str(RandomNumber))).Select
                CommandBars.ExecuteMso ("ShapesSubtract")
                
                With ActiveWindow.Selection.ShapeRange
                    .Name = "FivePointStar" + Str(StarsCount + 5) + Str(RandomNumber)
                End With
                
            End If
            
            ReDim Preserve StarsArray(UBound(StarsArray) + 1)
            StarsArray(UBound(StarsArray)) = "FivePointStar" + Str(StarsCount + 5) + Str(RandomNumber)
            
        End If
        
    Next
    
    ActiveWindow.Selection.SlideRange(1).Shapes.Range(StarsArray).Group
    
End Sub
