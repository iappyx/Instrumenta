Attribute VB_Name = "ModuleRAGStatus"
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

Sub GenerateRAGStatus(RAGColor As String)
    
    Set myDocument = Application.ActiveWindow
    Dim RAGStatus As Object
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
            
            If LCase(RAGColor) = "green" Then
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "GreenStatus" + Str(RandomNumber)
        End With
    
        Set AmberStatus = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 134, 104, 26, 26)
        
        With AmberStatus
            .Line.Visible = False

            If LCase(RAGColor) = "amber" Then
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "AmberStatus" + Str(RandomNumber)
        End With
    
        Set RedStatus = myDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 164, 104, 26, 26)
        
        With RedStatus
            .Line.Visible = False
            
            If LCase(RAGColor) = "red" Then
            .Fill.ForeColor.RGB = RGB(192, 0, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "RedStatus" + Str(RandomNumber)
        End With
        
        Set RAGStatus = ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("RAGBackground" + Str(RandomNumber), "GreenStatus" + Str(RandomNumber), "AmberStatus" + Str(RandomNumber), "RedStatus" + Str(RandomNumber))).Group
        RAGStatus.Name = "RAGStatus" + Str(RandomNumber)
    
End Sub
