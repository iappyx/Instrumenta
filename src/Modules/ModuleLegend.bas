Attribute VB_Name = "ModuleLegend"
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

Sub ShowCustomInsertLegend()

InsertLegendForm.Show

End Sub

Sub InsertLegendCustom()

Call InsertLegend(InsertLegendForm.LegendShapeComboBox.Value, InsertLegendForm.LegendOrientationComboBox.Value, InsertLegendForm.LegendNumberOfComboBox.ListIndex + 1)

End Sub

Sub InsertLegendSquareVerticalThree()
Call InsertLegend("square", "vertical", 3)
End Sub

Sub InsertLegendSquareHorizontalThree()
Call InsertLegend("square", "horizontal", 3)
End Sub

Sub InsertLegendCircleVerticalThree()
Call InsertLegend("circle", "vertical", 3)
End Sub

Sub InsertLegendCircleHorizontalThree()
Call InsertLegend("circle", "horizontal", 3)
End Sub

Sub InsertLegendSquareVerticalFive()
Call InsertLegend("square", "vertical", 5)
End Sub

Sub InsertLegendSquareHorizontalFive()
Call InsertLegend("square", "horizontal", 5)
End Sub

Sub InsertLegendCircleVerticalFive()
Call InsertLegend("circle", "vertical", 5)
End Sub

Sub InsertLegendCircleHorizontalFive()
Call InsertLegend("circle", "horizontal", 5)
End Sub

Sub InsertLegendSquareVerticalTen()
Call InsertLegend("square", "vertical", 10)
End Sub

Sub InsertLegendSquareHorizontalTen()
Call InsertLegend("square", "horizontal", 10)
End Sub

Sub InsertLegendCircleVerticalTen()
Call InsertLegend("circle", "vertical", 10)
End Sub

Sub InsertLegendCircleHorizontalTen()
Call InsertLegend("circle", "horizontal", 10)
End Sub


Function InsertLegend(shapeType As String, orientation As String, numberofitems As Long)

    Dim MyDocument As DocumentWindow
    Dim sld As Slide
    Dim msoType As MsoAutoShapeType
    Dim RandomNumber As Long
    
    Dim shp As shape
    Dim txt As shape
    
    Dim idx As Long
    Dim i As Long
    
    Dim leftPos As Single, topPos As Single, spacing As Single
    
    Set MyDocument = Application.ActiveWindow
    Set sld = MyDocument.View.Slide
    
    shapeType = LCase(Trim(shapeType))
    orientation = LCase(Trim(orientation))
    
    If shapeType = "circle" Then
        msoType = msoShapeOval
    ElseIf shapeType = "square" Then
        msoType = msoShapeRectangle
    Else
        MsgBox "Invalid shapeType. Use 'circle' or 'square'."
        Exit Function
    End If
    
    leftPos = 80
    topPos = 120
    
    Randomize
    RandomNumber = CLng(Rnd() * 1000000)
    
    ReDim arrNames(0 To ((numberofitems * 2) - 1))
    idx = 0
    
    For i = 1 To numberofitems
    
        If orientation = "horizontal" Then
            spacing = 100
            Set shp = sld.Shapes.AddShape(msoType, leftPos + ((i - 1) * spacing), topPos, 15, 15)
        ElseIf orientation = "vertical" Then
            spacing = 30
            Set shp = sld.Shapes.AddShape(msoType, leftPos, topPos + ((i - 1) * spacing), 15, 15)
        Else
            MsgBox "Invalid orientation. Use 'horizontal' or 'vertical'."
            Exit Function
        End If
        
        shp.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1 + ((i - 1) Mod 6)
        shp.Line.visible = msoFalse
        
        shp.Name = "LegendIcon_" & i & "_" & RandomNumber
        arrNames(idx) = shp.Name
        idx = idx + 1
        
        If orientation = "horizontal" Then
            Set txt = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, shp.left + 20, shp.Top - 2, 80, 20)
        Else
            Set txt = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, shp.left + 20, shp.Top - 2, 80, 20)
        End If
        
        txt.TextFrame.textRange.Text = "Legend " & i
        txt.TextFrame.textRange.Font.Size = 10
        
        txt.Name = "LegendText_" & i & "_" & RandomNumber
        arrNames(idx) = txt.Name
        idx = idx + 1
        
    Next i
    
    MyDocument.View.GotoSlide sld.SlideIndex
    
    Set LegendGroup = MyDocument.Selection.SlideRange(1).Shapes.Range(arrNames).Group
    LegendGroup.Name = "LegendGroup_" & RandomNumber
    LegendGroup.Tags.Add "INSTRUMENTA LEGEND", shapeType & orientation
    
    LegendGroup.left = 20
    LegendGroup.Top = 20

End Function

