Attribute VB_Name = "ModuleObjectsText"
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

Sub ObjectsTextInsertSpecialCharacter(SpecialCharacter As Long)
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        
        Application.ActiveWindow.Selection.TextRange.InsertSymbol Application.ActiveWindow.Selection.TextRange.Font.Name, SpecialCharacter, MsoTriState.msoTrue
        
    End If
    
End Sub

Sub ObjectsIncreaseLineSpacing()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.HasTextFrame Then
            
            With myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat
                .SpaceWithin = .SpaceWithin + 0.1
            End With
            
        Else
            
            MsgBox "No Text capable shapes selected."
            
        End If
        
    End If
    
End Sub

Sub ObjectsDecreaseLineSpacing()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.HasTextFrame Then
            
            With myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat
                
                If .SpaceWithin <= 0.1 Then
                    .SpaceWithin = 0
                Else
                    .SpaceWithin = .SpaceWithin - 0.1
                End If
                
            End With
            
        Else
            
            MsgBox "No Text capable shapes selected."
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveText()
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsRemoveTextLoop myDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            ObjectsRemoveTextLoop myDocument.Selection.ShapeRange
            
        End If
        
    End If
    
End Sub

Sub ObjectsRemoveTextLoop(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsRemoveTextLoop SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            SlideShape.TextFrame.TextRange.Text = ""
            
        End If
        
    End If
    
End Sub

Sub ObjectsSwapTextNoFormatting()
    
    Dim text1, text2 As String
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.Count = 2 Then
            
            If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                text1 = myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text
                text2 = myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text
                myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
                myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        Else
            
            MsgBox "Select two shapes To swap their text."
            
        End If
        
    End If
    
End Sub

Sub ObjectsSwapText()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.Count = 2 Then
            
            If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then
                
                Dim SlidePlaceHolder As PowerPoint.Shape
                Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
                
                myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Cut
                SlidePlaceHolder.TextFrame.TextRange.Paste
                
                myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Cut
                myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Paste
                
                SlidePlaceHolder.TextFrame.TextRange.Cut
                myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Paste
                
                SlidePlaceHolder.Delete
                
            Else
                
                MsgBox "Select two shapes that (can) have text."
                
            End If
            
        Else
            
            MsgBox "Select two shapes To swap their text."
            
        End If
        
    End If
    
End Sub

Sub ObjectsMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.HasTextFrame Then
            
            With myDocument.Selection.ShapeRange.TextFrame
                .MarginBottom = 0
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                
            End With
            
        Else
            
            MsgBox "No Text capable shape selected."
            
        End If
        
    End If
    
End Sub

Sub ObjectsMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.HasTextFrame Then
            
            With myDocument.Selection.ShapeRange.TextFrame
                .MarginBottom = .MarginBottom + 0.2
                .MarginLeft = .MarginLeft + 0.2
                .MarginRight = .MarginRight + 0.2
                .MarginTop = .MarginTop + 0.2
                
            End With
            
        Else
            
            MsgBox "No Text capable shape selected."
            
        End If
        
    End If
End Sub

Sub ObjectsMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.ShapeRange.HasTextFrame Then
            
            With myDocument.Selection.ShapeRange.TextFrame
                If .MarginBottom >= 0.2 Then
                    .MarginBottom = .MarginBottom - 0.2
                End If
                If .MarginLeft >= 0.2 Then
                    .MarginLeft = .MarginLeft - 0.2
                End If
                If .MarginRight >= 0.2 Then
                    .MarginRight = .MarginRight - 0.2
                End If
                If .MarginTop >= 0.2 Then
                    .MarginTop = .MarginTop - 0.2
                End If
                
            End With
            
        Else
            
            MsgBox "No Text capable shape selected."
            
        End If
        
    End If
End Sub

Sub ObjectsTextWordwrapToggle()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsTextWordwrapToggleLoop myDocument.Selection.ChildShapeRange(i)
            Next i
            
        Else
            
            ObjectsTextWordwrapToggleLoop myDocument.Selection.ShapeRange
            
        End If
        
    End If
End Sub

Sub ObjectsTextWordwrapToggleLoop(SlideShape)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsTextWordwrapToggleLoop SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If SlideShape.TextFrame.WordWrap = True Then
                SlideShape.TextFrame.WordWrap = False
            Else
                SlideShape.TextFrame.WordWrap = True
            End If
            
        End If
        
    End If
    
End Sub

Sub ObjectsAutoSizeTextToFitShape()
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop myDocument.Selection.ChildShapeRange(i), 2
            Next i
            
        Else
            
            ObjectsToggleAutoSizeLoop myDocument.Selection.ShapeRange, 2
            
        End If
        
    End If
End Sub

Sub ObjectsAutoSizeShapeToFitText()
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop myDocument.Selection.ChildShapeRange(i), 1
            Next i
            
        Else
            
            ObjectsToggleAutoSizeLoop myDocument.Selection.ShapeRange, 1
            
        End If
        
    End If
End Sub

Sub ObjectsAutoSizeNone()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop myDocument.Selection.ChildShapeRange(i), 0
            Next i
            
        Else
            
            ObjectsToggleAutoSizeLoop myDocument.Selection.ShapeRange, 0
            
        End If
        
    End If
End Sub

Sub ObjectsToggleAutoSize()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
        
        If myDocument.Selection.HasChildShapeRange Then
            
            For i = 1 To myDocument.Selection.ChildShapeRange.Count
                ObjectsToggleAutoSizeLoop myDocument.Selection.ChildShapeRange(i), 5
            Next i
            
        Else
            
            ObjectsToggleAutoSizeLoop myDocument.Selection.ShapeRange, 5
            
        End If
        
    End If
End Sub

Sub ObjectsToggleAutoSizeLoop(SlideShape, AutoSizeNum)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ObjectsToggleAutoSizeLoop SlideShapeChild, AutoSizeNum
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If AutoSizeNum = 5 Then
                With SlideShape.TextFrame2
                    If .AutoSize = 0 Then
                        .AutoSize = 1
                    ElseIf .AutoSize = 1 Then
                        .AutoSize = 2
                    ElseIf .AutoSize = 2 Then
                        .AutoSize = 0
                    ElseIf .AutoSize = -2 Then
                        .AutoSize = 0
                    End If
                    
                End With
                
            Else
                
                SlideShape.TextFrame2.AutoSize = AutoSizeNum
                
            End If
            
        End If
        
    End If
    
End Sub
