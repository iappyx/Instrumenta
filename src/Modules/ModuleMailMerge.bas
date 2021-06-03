Attribute VB_Name = "ModuleMailMerge"
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

Public CancelTriggered As Boolean


Sub InsertMergeField()

If ActiveWindow.Selection.Type = ppSelectionText Then

Application.ActiveWindow.Selection.TextRange.InsertAfter ("{{field name}}")

End If

End Sub

Sub ExcelMailMerge()
    
    #If Mac Then
        MsgBox "This Function will not work on a Mac"
    #Else
    
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
    
    Dim ExcelFile   As String
    Dim SlideShape  As Shape
    
    Dim ExcelApplication, ExcelSourceSheet, ExcelSourceWorkbook As Object
    
    'Early binding equivalent for reference:
    'Dim ExcelApplication As Excel.Application
    'Dim ExcelSourceWorkbook As Workbook
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .Show
        
        If .SelectedItems.Count = 0 Then
            MsgBox "No file selected."
            Exit Sub
        Else
            ExcelFile = .SelectedItems.Item(1)
        End If
        
    End With
    
    On Error Resume Next
    Set ExcelApplication = GetObject(Class:="Excel.Application")
    Err.Clear
    If ExcelApplication Is Nothing Then Set ExcelApplication = CreateObject(Class:="Excel.Application")
    On Error GoTo 0
    
    'Early binding equivalent for reference:
    'Set ExcelApplication = New Excel.Application
    
    Set ExcelSourceWorkbook = ExcelApplication.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    Set ExcelSourceSheet = ExcelSourceWorkbook.Sheets(1)
    
    On Error GoTo HandleError
    Set LastCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=1, SearchDirection:=2).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=2, SearchDirection:=2).Column)
    Set FirstCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=1, SearchDirection:=1).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=2, SearchDirection:=1).Column)
    On Error GoTo 0
    
    'Early binding equivalent for reference:
    'Set LastCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, ExcelSourceSheet.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column)
    'Set FirstCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row, ExcelSourceSheet.Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column)
    
    Dim MergeFields() As Variant
    Dim MergeTexts() As Variant
    MergeFields = ExcelSourceSheet.Range(FirstCell.Address & ":" & ExcelSourceSheet.Cells(FirstCell.Row, LastCell.Column).Address).Value
    MergeTexts = ExcelSourceSheet.Range(ExcelSourceSheet.Cells(FirstCell.Row + 1, FirstCell.Column).Address & ":" & ExcelSourceSheet.Cells(LastCell.Row, LastCell.Column).Address).Value
    
    
    PreviewMailMerge.MailMergeHeadersListBox.Clear
    PreviewMailMerge.MailMergeHeadersListBox.ColumnCount = UBound(MergeFields, 2)
    PreviewMailMerge.MailMergeHeadersListBox.List = ExcelSourceSheet.Range(FirstCell.Address & ":" & ExcelSourceSheet.Cells(FirstCell.Row, LastCell.Column).Address).Value
    
    PreviewMailMerge.MailMergeListBox.Clear
    PreviewMailMerge.MailMergeListBox.ColumnCount = UBound(MergeFields, 2)
    PreviewMailMerge.MailMergeListBox.List = ExcelSourceSheet.Range(ExcelSourceSheet.Cells(FirstCell.Row + 1, FirstCell.Column).Address & ":" & LastCell.Address).Value
    
    PreviewMailMerge.ExampleLabel.Caption = "Data taken from the first sheet of the Excel-file. Current selected slide will be duplicated" & Str(UBound(MergeTexts, 1)) & " times and all mail merge fields placed between {{ }} will be replaced with the data above." & vbNewLine & vbNewLine & "Example: {{" & MergeFields(1, 1) & "}}" & " will be replaced with " & MergeTexts(1, 1) & " on the first slide."
    
    
    ExcelSourceWorkbook.Close
    
    CancelTriggered = False
    
    PreviewMailMerge.Show
    
    If CancelTriggered = True Then Exit Sub
    
    Dim TempMergeFields As Variant
    Dim TempMergeTexts As Variant
    
    ProgressForm.Show
    
    For i = LBound(MergeFields, 1) To UBound(MergeFields, 1)
    
    SetProgress (i / UBound(MergeFields, 1) * 100)
    
        For j = LBound(MergeFields, 2) To UBound(MergeFields, 2)
            If j = 1 Then
                TempMergeFields = Array(MergeFields(i, j))
                
            Else
                ReDim Preserve TempMergeFields(UBound(TempMergeFields) + 1)
            End If
            TempMergeFields(j - 1) = "{{" & MergeFields(i, j) & "}}"
        Next j
    Next i
    
    For i = UBound(MergeTexts, 1) To LBound(MergeTexts, 1) Step -1
   
    For j = LBound(MergeTexts, 2) To UBound(MergeTexts, 2)
        
        If i < UBound(MergeTexts, 1) Then
            TempMergeTexts(j - 1) = ""
        ElseIf j = 1 Then
            TempMergeTexts = Array("")
        Else
            ReDim Preserve TempMergeTexts(UBound(TempMergeTexts) + 1)
        End If
        TempMergeTexts(j - 1) = MergeTexts(i, j)
    Next j
    
           
    Set MailMergeSlide = ActiveWindow.Selection.SlideRange(1).Duplicate
        
    For Each SlideShape In MailMergeSlide.Shapes
        ReplaceMergeFields SlideShape, TempMergeFields, TempMergeTexts
    Next SlideShape
        
    Next i
    
    ProgressForm.Hide
    
    Else
    
    MsgBox "No slide selected." & vbNewLine & vbNewLine & "Please select a slide that contains the merge fields as {{fieldname}} in shapes, tables and SmartArt."
    
    End If
    
    Exit Sub
    
HandleError:
    ExcelSourceWorkbook.Close
    MsgBox "Cannot load data. Does the first sheet in the Excel-file contain data with headers?"
    
    #End If
    
End Sub

'Sub ManualMailMerge()
      
    'Dim SlideShape  As Shape
    
    'ProgressForm.Show
    
    'For Each PresentationSlide In ActivePresentation.Slides
        
    '    SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
    '    For Each SlideShape In PresentationSlide.Shapes
    '        ReplaceMergeFields SlideShape, MergeFields, MergeTexts
    '    Next SlideShape
        
    'Next PresentationSlide
    
    'ProgressForm.hide
    
'End Sub

Sub ReplaceMergeFields(SlideShape, MergeFields As Variant, MergeTexts As Variant)
    
    'Dim ShapeTextRange As TextRange
    'Dim TemporaryTextRange As TextRange
    Dim MergeFieldsCount As Long
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ReplaceMergeFields SlideShapeChild, MergeFields, MergeTexts
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If Not SlideShape.TextFrame.TextRange = "" Then
                
                For MergeFieldsCount = LBound(MergeFields) To UBound(MergeFields)
                    
                    Set ShapeTextRange = SlideShape.TextFrame.TextRange
                    Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                    
                    Do While Not TemporaryTextRange Is Nothing
                        Set ShapeTextRange = ShapeTextRange.Characters(TemporaryTextRange.Start + TemporaryTextRange.Length, ShapeTextRange.Length)
                        Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                    Loop
                    
                Next MergeFieldsCount
                
            End If
            
        End If
        
        If SlideShape.HasTable Then
            For TableRow = 1 To SlideShape.Table.Rows.Count
                For TableColumn = 1 To SlideShape.Table.Columns.Count
                    
                    If Not SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame.TextRange = "" Then
                        
                        For MergeFieldsCount = LBound(MergeFields) To UBound(MergeFields)
                            
                            Set ShapeTextRange = SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame.TextRange
                            Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                            
                            Do While Not TemporaryTextRange Is Nothing
                                Set ShapeTextRange = ShapeTextRange.Characters(TemporaryTextRange.Start + TemporaryTextRange.Length, ShapeTextRange.Length)
                                Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                            Loop
                            
                        Next MergeFieldsCount
                        
                    End If
                    
                Next
            Next
        End If
        
        If SlideShape.HasSmartArt Then
            
            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.Count
                
                For Each SlideSmartArtNode In SlideShape.SmartArt.AllNodes
                    
                    If Not SlideSmartArtNode.TextFrame2.TextRange = "" Then
                        
                        For MergeFieldsCount = LBound(MergeFields) To UBound(MergeFields)
                            
                            Set ShapeTextRange = SlideSmartArtNode.TextFrame2.TextRange
                            
                            Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                            
                            'Needs fix, currently only first match is found, has to do with textframe2/textrange2
                            'Do While Not TemporaryTextRange Is Nothing
                            '    Set ShapeTextRange = ShapeTextRange.Characters(TemporaryTextRange.Start + TemporaryTextRange.Length, ShapeTextRange.Length)
                            '    Set TemporaryTextRange = ShapeTextRange.Replace(FindWhat:=MergeFields(MergeFieldsCount), Replacewhat:=MergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                            'Loop
                            
                        Next MergeFieldsCount
                        
                    End If
                    
                Next
                
            Next
            
        End If
        
    End If
    
End Sub
