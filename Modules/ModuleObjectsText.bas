Attribute VB_Name = "ModuleObjectsText"
Sub ObjectsRemoveText()
    Set myDocument = Application.ActiveWindow
    myDocument.Selection.ShapeRange.TextFrame.TextRange.Text = ""
End Sub

Sub ObjectsSwapText()
    Dim text1, text2 As String
    Set myDocument = Application.ActiveWindow
    text1 = myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text
    text2 = myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1
End Sub

Sub ObjectsMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    With myDocument.Selection.ShapeRange.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        
    End With
    
End Sub

Sub ObjectsMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    With myDocument.Selection.ShapeRange.TextFrame
        .MarginBottom = .MarginBottom + 0.2
        .MarginLeft = .MarginLeft + 0.2
        .MarginRight = .MarginRight + 0.2
        .MarginTop = .MarginTop + 0.2
        
    End With
    
End Sub

Sub ObjectsMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
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
    
End Sub

Sub ObjectsTextWordwrapToggle()
    
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.ShapeRange.TextFrame.WordWrap = True Then
        myDocument.Selection.ShapeRange.TextFrame.WordWrap = False
    Else
        myDocument.Selection.ShapeRange.TextFrame.WordWrap = True
    End If
    
End Sub
