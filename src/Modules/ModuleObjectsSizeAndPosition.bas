Attribute VB_Name = "ModuleObjectsSizeAndPosition"
Sub ObjectsSizeToTallest()
    Set myDocument = Application.ActiveWindow
    Dim Tallest     As Single
    Tallest = myDocument.Selection.ShapeRange(1).Height
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Height > Tallest Then Tallest = SlideShape.Height
    Next
    
    myDocument.Selection.ShapeRange.Height = Tallest
    
End Sub

Sub ObjectsSizeToShortest()
    Set myDocument = Application.ActiveWindow
    Dim Shortest    As Single
    Shortest = myDocument.Selection.ShapeRange(1).Height
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Height < Shortest Then Shortest = SlideShape.Height
    Next
    
    myDocument.Selection.ShapeRange.Height = Shortest
    
End Sub

Sub ObjectsSizeToWidest()
    Set myDocument = Application.ActiveWindow
    Dim Widest      As Single
    Widest = myDocument.Selection.ShapeRange(1).Width
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Width > Widest Then Widest = SlideShape.Width
    Next
    
    myDocument.Selection.ShapeRange.Width = Widest
    
End Sub

Sub ObjectsSizeToNarrowest()
    Set myDocument = Application.ActiveWindow
    Dim Narrowest   As Single
    Narrowest = myDocument.Selection.ShapeRange(1).Width
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        If SlideShape.Width < Narrowest Then Narrowest = SlideShape.Width
    Next
    
    myDocument.Selection.ShapeRange.Width = Narrowest
    
End Sub

Sub ObjectsSameHeight()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
    
End Sub

Sub ObjectsSameWidth()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    
End Sub

Sub ObjectsSameHeightAndWidth()
    Set myDocument = Application.ActiveWindow
    
    myDocument.Selection.ShapeRange.Height = myDocument.Selection.ShapeRange(1).Height
    myDocument.Selection.ShapeRange.Width = myDocument.Selection.ShapeRange(1).Width
    
End Sub

Sub ObjectsSwapPosition()
    
    Dim Left1, Left2, Top1, Top2 As Single
    
    Left1 = ActiveWindow.Selection.ShapeRange(1).Left
    Left2 = ActiveWindow.Selection.ShapeRange(2).Left
    Top1 = ActiveWindow.Selection.ShapeRange(1).Top
    Top2 = ActiveWindow.Selection.ShapeRange(2).Top
    
    ActiveWindow.Selection.ShapeRange(1).Left = Left2
    ActiveWindow.Selection.ShapeRange(2).Left = Left1
    ActiveWindow.Selection.ShapeRange(1).Top = Top2
    ActiveWindow.Selection.ShapeRange(2).Top = Top1
    
End Sub
