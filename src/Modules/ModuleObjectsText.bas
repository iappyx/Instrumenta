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
