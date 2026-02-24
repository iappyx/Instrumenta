Attribute VB_Name = "InsertLegendForm"
Attribute VB_Base = "0{0D0051EB-3FCD-42B4-9CF5-DB353AAE20DD}{2E1AD606-929D-4827-8D13-03C5591B15D8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

Private Sub UserForm_Activate()

LegendNumberOfComboBox.Clear
LegendNumberOfComboBox.AddItem ("1")
LegendNumberOfComboBox.AddItem ("2")
LegendNumberOfComboBox.AddItem ("3")
LegendNumberOfComboBox.AddItem ("4")
LegendNumberOfComboBox.AddItem ("5")
LegendNumberOfComboBox.AddItem ("6")
LegendNumberOfComboBox.AddItem ("7")
LegendNumberOfComboBox.AddItem ("8")
LegendNumberOfComboBox.AddItem ("9")
LegendNumberOfComboBox.AddItem ("10")
LegendNumberOfComboBox.AddItem ("11")
LegendNumberOfComboBox.AddItem ("12")
LegendNumberOfComboBox.AddItem ("13")
LegendNumberOfComboBox.AddItem ("14")
LegendNumberOfComboBox.AddItem ("15")
LegendNumberOfComboBox.AddItem ("16")
LegendNumberOfComboBox.AddItem ("17")
LegendNumberOfComboBox.AddItem ("18")
LegendNumberOfComboBox.AddItem ("19")
LegendNumberOfComboBox.AddItem ("20")

LegendShapeComboBox.Clear
LegendShapeComboBox.AddItem ("circle")
LegendShapeComboBox.AddItem ("square")

LegendOrientationComboBox.Clear
LegendOrientationComboBox.AddItem ("horizontal")
LegendOrientationComboBox.AddItem ("vertical")

LegendNumberOfComboBox.ListIndex = 0
LegendShapeComboBox.ListIndex = 0
LegendOrientationComboBox.ListIndex = 0

End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub InsertLegendsButton_Click()
    InsertLegendCustom
    Me.Hide
End Sub
