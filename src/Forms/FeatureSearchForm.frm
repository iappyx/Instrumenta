VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FeatureSearchForm 
   Caption         =   "Search Instrumenta features"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   OleObjectBlob   =   "FeatureSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FeatureSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private SearchResultIndices As String


Private Sub UserForm_Initialize()
   
    With lstResults
        .ColumnCount = 3
        .ColumnWidths = "200;225;225"
    End With
    
    
    PerformSearch ""
End Sub

Private Sub txtSearch_Change()
    PerformSearch txtSearch.Text
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lstResults.ListCount > 0 Then
                lstResults.SetFocus
                If lstResults.ListIndex = -1 Then lstResults.ListIndex = 0
            End If
            KeyCode = 0
        Case vbKeyReturn
            If lstResults.ListCount > 0 Then
                If lstResults.ListIndex = -1 Then lstResults.ListIndex = 0
                ExecuteSelectedFeature
            End If
            KeyCode = 0
    End Select
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ExecuteSelectedFeature
End Sub

Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ExecuteSelectedFeature
        KeyCode = 0
    End If
End Sub

Private Sub cmdExecute_Click()
    ExecuteSelectedFeature
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PerformSearch(query As String)
    Dim indices() As String
    Dim i As Long
    Dim feat As FeatureData
    
    lstResults.Clear
    
    SearchResultIndices = ModuleFeatureSearch.SearchFeatures(query)
    
    If Len(SearchResultIndices) = 0 Then
        lblResultCount.Caption = "0 features found"
        Exit Sub
    End If
    
    indices = Split(left(SearchResultIndices, Len(SearchResultIndices) - 1), "|")
    
    For i = LBound(indices) To UBound(indices)
        feat = ModuleFeatureSearch.GetFeatureByIndex(CLng(indices(i)))
        
        lstResults.AddItem feat.Label
        lstResults.List(lstResults.ListCount - 1, 1) = feat.TabSingleView
        lstResults.List(lstResults.ListCount - 1, 2) = feat.TabMultiView
    Next i

    If UBound(indices) - LBound(indices) + 1 = 1 Then
        lblResultCount.Caption = "1 feature found"
    Else
        lblResultCount.Caption = (UBound(indices) - LBound(indices) + 1) & " features found"
    End If
    
    If lstResults.ListCount > 0 Then lstResults.ListIndex = 0
End Sub

Private Sub ExecuteSelectedFeature()
    Dim selectedIndex As Long
    Dim indices() As String
    Dim feat As FeatureData
    
    selectedIndex = lstResults.ListIndex
    
    If selectedIndex < 0 Or Len(SearchResultIndices) = 0 Then Exit Sub
    
    indices = Split(left(SearchResultIndices, Len(SearchResultIndices) - 1), "|")
    
    If selectedIndex <= UBound(indices) Then
        feat = ModuleFeatureSearch.GetFeatureByIndex(CLng(indices(selectedIndex)))
        
        ModuleFeatureSearch.ExecuteFeature feat.OnAction
        txtSearch.SetFocus
    End If
End Sub

