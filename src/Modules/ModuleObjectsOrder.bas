
Attribute VB_Name = "ModuleObjectsOrder"
Option Explicit

' Moves the selected object one layer forward
Sub BringForward()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.ZOrder msoBringForward
    Else
        MsgBox "Please select an object.", vbExclamation
    End If
End Sub

' Brings the selected object to the front
Sub BringToFront()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.ZOrder msoBringToFront
    Else
        MsgBox "Please select an object.", vbExclamation
    End If
End Sub

' Moves the selected object one layer backward
Sub SendBackward()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.ZOrder msoSendBackward
    Else
        MsgBox "Please select an object.", vbExclamation
    End If
End Sub

' Sends the selected object to the back
Sub SendToBack()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.ZOrder msoSendToBack
    Else
        MsgBox "Please select an object.", vbExclamation
    End If
End Sub
