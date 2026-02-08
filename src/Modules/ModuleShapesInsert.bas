Attribute VB_Name = "ModuleShapesInsert"
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


Function ShapesInsertRectangle()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRectangle, 100, 100, 150, 100
End Function

Function ShapesInsertRoundedRectangle()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRoundedRectangle, 100, 100, 150, 100
End Function

Function ShapesInsertOval()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeOval, 100, 100, 150, 100
End Function

Function ShapesInsertTriangle()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeIsoscelesTriangle, 100, 100, 150, 100
End Function

Function ShapesInsertRightTriangle()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRightTriangle, 100, 100, 150, 100
End Function

Function ShapesInsertParallelogram()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeParallelogram, 100, 100, 150, 100
End Function

Function ShapesInsertTrapezoid()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeTrapezoid, 100, 100, 150, 100
End Function

Function ShapesInsertPentagon()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapePentagon, 100, 100, 120, 120
End Function

Function ShapesInsertHexagon()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeHexagon, 100, 100, 120, 120
End Function

Function ShapesInsertOctagon()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeOctagon, 100, 100, 120, 120
End Function

Function ShapesInsertStraightLine()
    Application.ActiveWindow.View.Slide.Shapes.AddLine 100, 100, 250, 100
End Function

Function ShapesInsertStraightArrow()
    With Application.ActiveWindow.View.Slide.Shapes.AddLine(100, 100, 250, 100)
        .Line.EndArrowheadStyle = msoArrowheadOpen
    End With
End Function

Function ShapesInsertRightArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRightArrow, 100, 100, 150, 50
End Function

Function ShapesInsertLeftArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeLeftArrow, 100, 100, 150, 50
End Function

Function ShapesInsertUpArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeUpArrow, 100, 100, 50, 150
End Function

Function ShapesInsertDownArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeDownArrow, 100, 100, 50, 150
End Function

Function ShapesInsertCurvedRightArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeCurvedRightArrow, 100, 100, 150, 100
End Function

Function ShapesInsertBentArrow()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeBentArrow, 100, 100, 100, 100
End Function

Function ShapesInsertRoundedRectangularCallout()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRoundedRectangularCallout, 100, 100, 200, 100
End Function

Function ShapesInsertCloudCallout()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeCloudCallout, 100, 100, 200, 100
End Function

Function ShapesInsertOvalCallout()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeOvalCallout, 100, 100, 200, 100
End Function

Function ShapesInsertFlowchartProcess()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeFlowchartProcess, 100, 100, 150, 100
End Function

Function ShapesInsertFlowchartDecision()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeFlowchartDecision, 100, 100, 150, 150
End Function

Function ShapesInsertFlowchartTerminator()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeFlowchartTerminator, 100, 100, 150, 50
End Function

Function ShapesInsertFlowchartConnector()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeFlowchartConnector, 100, 100, 50, 50
End Function

Function ShapesInsertStar4()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape4pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar5()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape5pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar6()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape6pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar7()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape7pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar8()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape8pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar10()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape10pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar12()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape12pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar16()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape16pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar24()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape24pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertStar32()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShape32pointStar, 100, 100, 100, 100
End Function

Function ShapesInsertWave()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeWave, 100, 100, 200, 50
End Function

Function ShapesInsertRightBrace()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRightBrace, 100, 100, 50, 150
End Function

Function ShapesInsertLeftBrace()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeLeftBrace, 100, 100, 50, 150
End Function

Function ShapesInsertRightBracket()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeRightBracket, 100, 100, 50, 150
End Function

Function ShapesInsertLeftBracket()
    Application.ActiveWindow.View.Slide.Shapes.AddShape msoShapeLeftBracket, 100, 100, 50, 150
End Function

