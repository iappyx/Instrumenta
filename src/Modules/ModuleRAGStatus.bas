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

Sub AverageRAGStatus()

    Set MyDocument = Application.ActiveWindow
    Dim RAGStatusCount As Integer
    Dim RAGStatusSum As Double
    
    RAGStatusSum = 0
    RAGStatusCount = 0
          
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each shape In ActiveWindow.Selection.ShapeRange
            
            If (InStr(shape.Name, "RAGStatus") = 1) And (Not shape.Tags("INSTRUMENTA RAGSTATUS") = "") Then
                
                RAGStatusCount = RAGStatusCount + 1
                
                If shape.Tags("INSTRUMENTA RAGSTATUS") = "green" Then
                    RAGStatusSum = RAGStatusSum + 3
                ElseIf shape.Tags("INSTRUMENTA RAGSTATUS") = "amber" Then
                    RAGStatusSum = RAGStatusSum + 6
                ElseIf shape.Tags("INSTRUMENTA RAGSTATUS") = "red" Then
                    RAGStatusSum = RAGStatusSum + 9
                End If
                
            End If
            
        Next shape
    End If
    
    If RAGStatusCount > 0 Then
    
    ActiveWindow.Selection.Unselect
    
    If Round((RAGStatusSum / RAGStatusCount) / 3, 0) * 3 = 3 Then
    
    GenerateRAGStatus "green"
    
    ElseIf Round((RAGStatusSum / RAGStatusCount) / 3, 0) * 3 = 6 Then
    
    GenerateRAGStatus "amber"
    
    ElseIf Round((RAGStatusSum / RAGStatusCount) / 3, 0) * 3 = 9 Then
    
    GenerateRAGStatus "red"
    
    End If
    
    Else
    MsgBox "No RAG status shape selected."
    End If

End Sub

Sub GenerateRAGStatus(RAGColor As String)
    
    Set MyDocument = Application.ActiveWindow
    
    Dim ExistingWidth, ExistingHeight, ExistingTop, ExistingLeft, ExistingRotation As Double
    Dim ExistingRAGStatus As Boolean
    ExistingRAGStatus = False
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        For Each shape In ActiveWindow.Selection.ShapeRange
            
            If InStr(shape.Name, "RAGStatus") = 1 Then
                
                ExistingRAGStatus = True
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
    
    
    Dim RAGStatus As Object
    RandomNumber = Round(Rnd() * 1000000, 0)
    
        Set RAGBackground = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeRoundedRectangle, 100, 100, 94, 34)
        
        With RAGBackground
            .Line.visible = False
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Name = "RAGBackground" + Str(RandomNumber)
        End With
        
        
        Set GreenStatus = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 104, 104, 26, 26)
        
        With GreenStatus
            .Line.visible = False
            
            If LCase(RAGColor) = "green" Then
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "GreenStatus" + Str(RandomNumber)
        End With
    
        Set AmberStatus = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 134, 104, 26, 26)
        
        With AmberStatus
            .Line.visible = False

            If LCase(RAGColor) = "amber" Then
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "AmberStatus" + Str(RandomNumber)
        End With
    
        Set RedStatus = MyDocument.Selection.SlideRange.Shapes.AddShape(msoShapeOval, 164, 104, 26, 26)
        
        With RedStatus
            .Line.visible = False
            
            If LCase(RAGColor) = "red" Then
            .Fill.ForeColor.RGB = RGB(192, 0, 0)
            Else
            .Fill.ForeColor.RGB = RGB(59, 56, 56)
            End If
            
            .Name = "RedStatus" + Str(RandomNumber)
        End With
        
        Set RAGStatus = ActiveWindow.Selection.SlideRange(1).Shapes.Range(Array("RAGBackground" + Str(RandomNumber), "GreenStatus" + Str(RandomNumber), "AmberStatus" + Str(RandomNumber), "RedStatus" + Str(RandomNumber))).Group
        RAGStatus.Name = "RAGStatus" + Str(RandomNumber)
        RAGStatus.Tags.Add "INSTRUMENTA RAGSTATUS", RAGColor
        
        If ExistingRAGStatus = True Then
            RAGStatus.Width = ExistingWidth
            RAGStatus.Height = ExistingHeight
            RAGStatus.Top = ExistingTop
            RAGStatus.left = ExistingLeft
            RAGStatus.Rotation = ExistingRotation
        End If
    
End Sub
