Attribute VB_Name = "ModuleQRCodes"
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


Sub InsertQRCode()

Set MyDocument = Application.ActiveWindow
QRCodeText = ""

QRCodeText = InputBox("Please note that this functionality uses external APIs to generate the QR-code (goqr.me/api/)." & vbNewLine & vbNewLine & vbNewLine & "Please provide the URL (or other content) for the QR-code:", "Generate QR-code", "https://")

If Not QRCodeText = "" Then
    QRCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=" & QRCodeText
    Dim QRCode As shape
    Set QRCode = MyDocument.Selection.SlideRange.Shapes.AddPicture(QRCodeUrl, msoTrue, msoTrue, 0, 0)
    QRCode.Select
    ObjectsAlignCenters
    ObjectsAlignMiddles
End If

End Sub

