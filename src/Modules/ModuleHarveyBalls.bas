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

    Set MyDocument = Application.ActiveWindow
    Dim HarveyCount As Integer
    Dim HarveySum As Double
    
    HarveySum = 0
    HarveyCount = 0
          
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each shape In ActiveWindow.Selection.ShapeRange
            
            If (InStr(shape.Name, "HarveyBall") = 1) And (Not shape.Tags("INSTRUMENTA HARVEYBALL") = "") Then
                
                HarveyCount = HarveyCount + 1
                HarveySum = HarveySum + CDbl(shape.Tags("INSTRUMENTA HARVEYBALL"))
                
            End If
            
        Next shape
    End If
    
    If HarveyCount > 0 Then
    
    ActiveWindow.Selection.Unselect
    GenerateHarveyBallPercent Round(HarveySum / HarveyCount)
    
    Else
    MsgBox "No Harveyball-shapes selected."
    End If

End Sub

Sub GenerateHarveyBallPercent(FillPercentage As Double)
    Set MyDocument = Application.ActiveWindow
    
    Dim ExistingWidth, ExistingHeight, ExistingTop, ExistingLeft, ExistingRotation As Double
    Dim ExistingHarveyBall As Boolean
    ExistingHarveyBall = False
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each shape In ActiveWindow.Selection.ShapeRange
            
            If InStr(shape.Name, "HarveyBall") = 1 Then
                
                ExistingHarveyBall = True
                ExistingWidth = shape.Width
                ExistingHeight = shape.Height
                ExistingTop = shape.Top
                ExistingLeft = shape.left
                ExistingRotation = shape.Rotation
                shape.Delete
                
            End If
            
            Exit For
        Next shape
    End If
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim HarveyCircle, HarveyFill As shape
    
    Set HarveyCircle = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50)
    Set HarveyFill = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapePie, 101, 101, 48, 48)
    With HarveyFill
        .Name = "HarveyFill" + Str(RandomNumber)
        .Adjustments.Item(2) = -90
        .Adjustments.Item(1) = ((FillPercentage / 100) * 360) - 90
        .Line.visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    With HarveyCircle
        .Name = "HarveyCircle" + Str(RandomNumber)
        .Line.visible = False
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    If FillPercentage > 0 Then
        HarveyFill.Adjustments(1) = HarveyFill.Adjustments(1) - 0.1
    End If
    
    ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("HarveyCircle" + Str(RandomNumber), "HarveyFill" + Str(RandomNumber))).Select
    CommandBars.ExecuteMso ("ShapesCombine")
    For Each shape In ActiveWindow.Selection.ShapeRange
        
        shape.Name = "HarveyBall" + Str(RandomNumber)
        
        If ExistingHarveyBall = True Then
            shape.Width = ExistingWidth
            shape.Height = ExistingHeight
            shape.Top = ExistingTop
            shape.left = ExistingLeft
            shape.Rotation = ExistingRotation
        End If
        
        shape.Tags.Add "INSTRUMENTA HARVEYBALL", FillPercentage
        
    Next
    ActiveWindow.Selection.Unselect
    
End Sub

