Attribute VB_Name = "ModuleProcess"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

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

Sub InsertProcessSmartArt()

Set MyDocument = Application.ActiveWindow
Set ProcessSmartArt = MyDocument.Selection.SlideRange.shapes.AddSmartArt(Application.SmartArtLayouts("urn:microsoft.com/office/officeart/2005/8/layout/hChevron3"), 50, 100, Application.ActivePresentation.PageSetup.slideWidth - 100, 50)

For nodeCount = 1 To ProcessSmartArt.SmartArt.AllNodes.count

With ProcessSmartArt.SmartArt.AllNodes(nodeCount).TextFrame2.textRange
.Font.Size = 14
.Font.Bold = msoTrue
.text = "Step" & Str(nodeCount)
End With

Next

End Sub
   
    
    
    
