Attribute VB_Name = "ModuleFivePointStars"
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

Sub AverageFivePointStars()

    Set myDocument = Application.ActiveWindow
    Dim FivePointStarCount As Integer
    Dim FivePointStarSum As Double
    
    FivePointStarSum = 0
    FivePointStarCount = 0
          
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        For Each Shape In ActiveWindow.Selection.ShapeRange
            
            If (InStr(Shape.Name, "StarRating") = 1) And (Not Shape.Tags("INSTRUMENTA STARRATING") = "") Then
                
                FivePointStarCount = FivePointStarCount + 1
                FivePointStarSum = FivePointStarSum + CDbl(Shape.Tags("INSTRUMENTA STARRATING"))
                
            End If
            
        Next Shape
    End If
    
    If FivePointStarCount > 0 Then
    
    ActiveWindow.Selection.Unselect
    GenerateFivePointStars Round((FivePointStarSum / FivePointStarCount) / 0.5, 0) * 0.5
    
    Else
    MsgBox "No star rating shape selected."
    End If

End Sub


Sub GenerateFivePointStars(NumberOfStars As Double)
    
    Set myDocument = Application.ActiveWindow
    
    Dim ExistingWidth, ExistingHeight, ExistingTop, ExistingLeft, ExistingRotation As Double
    Dim ExistingStarRating As Boolean
    ExistingStarRating = False
    
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        For Each Shape In ActiveWindow.Selection.ShapeRange
            
            If InStr(Shape.Name, "StarRating") = 1 Then
                
                ExistingStarRating = True
                ExistingWidth = Shape.Width
                ExistingHeight = Shape.Height
                ExistingTop = Shape.Top
                ExistingLeft = Shape.Left
                ExistingRotation = Shape.Rotation
                Shape.Delete
                
            End If
            
            Exit For
        Next Shape
    End If
    
    
    Dim StarsCount  As Double
    Dim StarsArray  As Variant
    Dim StarRating As Object
    
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
    
    Set StarRating = ActiveWindow.Selection.SlideRange(1).Shapes.Range(StarsArray).Group
    StarRating.Name = "StarRating" + Str(RandomNumber)
    StarRating.Tags.Add "INSTRUMENTA STARRATING", NumberOfStars
    
    If ExistingStarRating = True Then
        StarRating.Width = ExistingWidth
        StarRating.Height = ExistingHeight
        StarRating.Top = ExistingTop
        StarRating.Left = ExistingLeft
        StarRating.Rotation = ExistingRotation
    End If
    
End Sub
