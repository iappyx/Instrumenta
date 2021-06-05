Attribute VB_Name = "ModuleHarveyBalls"
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

Sub AverageHarveyBall()

    Set myDocument = Application.ActiveWindow
    Dim HarveyCount As Integer
    Dim HarveySum As Double
    
    HarveySum = 0
    HarveyCount = 0
          
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        For Each Shape In ActiveWindow.Selection.ShapeRange
            
            If (InStr(Shape.Name, "HarveyBall") = 1) And (Not Shape.Tags("INSTRUMENTA HARVEYBALL") = "") Then
                
                HarveyCount = HarveyCount + 1
                HarveySum = HarveySum + CDbl(Shape.Tags("INSTRUMENTA HARVEYBALL"))
                
            End If
            
        Next Shape
    End If
    
    If HarveyCount > 0 Then
    
    ActiveWindow.Selection.Unselect
    GenerateHarveyBallPercent Round(HarveySum / HarveyCount)
    
    Else
    MsgBox "No Harveyball-shapes selected."
    End If

End Sub

Sub GenerateHarveyBallPercent(FillPercentage As Double)
    Set myDocument = Application.ActiveWindow
    
    Dim ExistingWidth, ExistingHeight, ExistingTop, ExistingLeft, ExistingRotation As Double
    Dim ExistingHarveyBall As Boolean
    ExistingHarveyBall = False
    
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        For Each Shape In ActiveWindow.Selection.ShapeRange
            
            If InStr(Shape.Name, "HarveyBall") = 1 Then
                
                ExistingHarveyBall = True
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
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim HarveyCircle, HarveyFill As Shape
    
    Set HarveyCircle = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50)
    Set HarveyFill = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapePie, 101, 101, 48, 48)
    With HarveyFill
        .Name = "HarveyFill" + Str(RandomNumber)
        .Adjustments.Item(2) = -90
        .Adjustments.Item(1) = ((FillPercentage / 100) * 360) - 90
        .Line.Visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    With HarveyCircle
        .Name = "HarveyCircle" + Str(RandomNumber)
        .Line.Visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    If FillPercentage > 0 Then
        HarveyFill.Adjustments(1) = HarveyFill.Adjustments(1) - 0.1
    End If
    
    ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("HarveyCircle" + Str(RandomNumber), "HarveyFill" + Str(RandomNumber))).Select
    CommandBars.ExecuteMso ("ShapesCombine")
    For Each Shape In ActiveWindow.Selection.ShapeRange
        
        Shape.Name = "HarveyBall" + Str(RandomNumber)
        
        If ExistingHarveyBall = True Then
            Shape.Width = ExistingWidth
            Shape.Height = ExistingHeight
            Shape.Top = ExistingTop
            Shape.Left = ExistingLeft
            Shape.Rotation = ExistingRotation
        End If
        
        Shape.Tags.Add "INSTRUMENTA HARVEYBALL", FillPercentage
        
    Next
    ActiveWindow.Selection.Unselect
    
End Sub

