Attribute VB_Name = "ModuleTicksAndCrosses"
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

Sub TextBulletsTicks()
 
    Set MyDocument = Application.ActiveWindow
     
    On Error Resume Next
    With MyDocument.Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 252
        .visible = True
        .Font.Name = "Wingdings"
        .Font.Color = RGB(0, 128, 0)
        
    End With
    On Error GoTo 0
    
End Sub

Sub TextBulletsCrosses()
    
    Set MyDocument = Application.ActiveWindow
   
    On Error Resume Next
    With MyDocument.Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 215
        .visible = True
        .Font.Name = "Calibri"
        .Font.Color = RGB(255, 0, 0)
        
    End With
    On Error GoTo 0
    
End Sub
