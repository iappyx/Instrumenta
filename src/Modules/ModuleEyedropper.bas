Attribute VB_Name = "ModuleEyedropper"
'================================================================================
' VBA Module: Eyedropper Simulator using Windows API
' Purpose: Reads the color of the pixel at the current mouse position and
'          applies it to the Fill or Outline property of all selected shapes.
'================================================================================

' --- 1. Windows API Declarations and Structure ---

#If VBA7 And Win64 Then
    ' 64-bit Declarations
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
    ' 32-bit Declarations
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

' Structure to store the mouse coordinates (x, y)
Private Type POINTAPI
    x As Long
    y As Long
End Type

' --- Helper Function: GetPixelColor ---

#If VBA7 And Win64 Then
    Private Function GetPixelColor() As Long
        Dim pt As POINTAPI
        Dim hdc As LongPtr
        Dim lColorRGB As Long
        
        GetCursorPos pt
        hdc = GetDC(0)
        lColorRGB = GetPixel(hdc, pt.x, pt.y)
        ReleaseDC 0, hdc
        
        GetPixelColor = lColorRGB
    End Function
#Else
    Private Function GetPixelColor() As Long
        Dim pt As POINTAPI
        Dim hdc As Long
        Dim lColorRGB As Long
        
        GetCursorPos pt
        hdc = GetDC(0)
        lColorRGB = GetPixel(hdc, pt.x, pt.y)
        ReleaseDC 0, hdc
        
        GetPixelColor = lColorRGB
    End Function
#End If

' --- Apply Pixel Color to FILL ---

Public Sub ApplyPixelColorToFill()
    Dim oSh As Shape
    Dim lColor As Long
    
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes before running the macro.", vbExclamation
        Exit Sub
    End If
    
    lColor = GetPixelColor()
    
    For Each oSh In ActiveWindow.Selection.ShapeRange
        ' Apply the captured color to the shape's fill
        On Error Resume Next
        oSh.Fill.ForeColor.RGB = lColor
        On Error GoTo 0
    Next oSh
End Sub

' --- Apply Pixel Color to OUTLINE (Line) ---

Public Sub ApplyPixelColorToOutline()
    Dim oSh As Shape
    Dim lColor As Long
    
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes before running the macro.", vbExclamation
        Exit Sub
    End If
    
    lColor = GetPixelColor()
    
    For Each oSh In ActiveWindow.Selection.ShapeRange
        ' Apply the captured color to the shape's line/outline
        On Error Resume Next
        
        ' Ensure the line is visible before setting the color
        oSh.Line.Visible = msoTrue
        oSh.Line.ForeColor.RGB = lColor
        
        On Error GoTo 0
    Next oSh
End Sub
