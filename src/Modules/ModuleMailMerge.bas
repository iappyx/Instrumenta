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
Public ManualHeaders() As Variant
Public ManualTexts() As Variant


Sub InsertMergeField()

If ActiveWindow.Selection.Type = ppSelectionText Then

Application.ActiveWindow.Selection.TextRange.InsertAfter ("{{fieldName}}")

End If

End Sub

Sub ImportHeadersFromExcel()

   
   If ActiveWindow.Selection.Type = ppSelectionText Then
    
    Dim ExcelFile   As String
    Dim SlideShape  As Shape
    
    Dim ExcelApplication, ExcelSourceSheet, ExcelSourceWorkbook As Object
    
    'Early binding equivalent for reference:
    'Dim ExcelApplication As Excel.Application
    'Dim ExcelSourceWorkbook As Workbook
    
    #If Mac Then
    
    ExcelFile = MacFileDialog("/")
    
    If ExcelFile = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    
    #Else
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
    #End If
    
    
    On Error Resume Next
    Set ExcelApplication = GetObject(Class:="Excel.Application")
    Err.Clear
    If ExcelApplication Is Nothing Then Set ExcelApplication = CreateObject(Class:="Excel.Application")
    On Error GoTo 0
    
    'Early binding equivalent for reference:
    'Set ExcelApplication = New Excel.Application
    
    #If Mac Then
    Set ExcelSourceWorkbook = ExcelApplication.Application.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #Else
    Set ExcelSourceWorkbook = ExcelApplication.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #End If
    Set ExcelSourceSheet = ExcelSourceWorkbook.Sheets(1)
    
    On Error GoTo HandleError
    Set LastCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=1, SearchDirection:=2).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=2, SearchDirection:=2).Column)
    Set FirstCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=1, SearchDirection:=1).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=2, SearchDirection:=1).Column)
    On Error GoTo 0
    
    'Early binding equivalent for reference:
    'Set LastCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, ExcelSourceSheet.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column)
    'Set FirstCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row, ExcelSourceSheet.Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column)
    
    Dim MergeFields() As Variant
    MergeFields = ExcelSourceSheet.Range(FirstCell.Address & ":" & ExcelSourceSheet.Cells(FirstCell.Row, LastCell.Column).Address).Value
    ExcelSourceWorkbook.Close

        
    For i = LBound(MergeFields, 1) To UBound(MergeFields, 1)
    
    SetProgress (i / UBound(MergeFields, 1) * 100)
    
        For j = LBound(MergeFields, 2) To UBound(MergeFields, 2)
            Application.ActiveWindow.Selection.TextRange.InsertAfter (" {{" & MergeFields(i, j) & "}}")
        Next j
    Next i

    Else
            
    MsgBox "Please select a shape or table where you want to paste the merge fields."
            
    End If

    
    Exit Sub
    
HandleError:
    ExcelSourceWorkbook.Close
    MsgBox "Cannot load data. Does the first sheet in the Excel-file contain data with headers?"
    

End Sub

Sub ExcelMailMerge()
    
   
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
    
    Dim ExcelFile   As String
    Dim SlideShape  As Shape
    
    Dim ExcelApplication, ExcelSourceSheet, ExcelSourceWorkbook As Object
    
    MailMergeSlideNum = ActiveWindow.Selection.SlideRange(1).SlideNumber
    
    
    'Early binding equivalent for reference:
    'Dim ExcelApplication As Excel.Application
    'Dim ExcelSourceWorkbook As Workbook
    
    #If Mac Then
    
    ExcelFile = MacFileDialog("/")
    
    If ExcelFile = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    
    #Else
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
    #End If
    
    On Error Resume Next
    Set ExcelApplication = GetObject(Class:="Excel.Application")
    Err.Clear
    If ExcelApplication Is Nothing Then Set ExcelApplication = CreateObject(Class:="Excel.Application")
    On Error GoTo 0
    
    'Early binding equivalent for reference:
    'Set ExcelApplication = New Excel.Application
    
    #If Mac Then
    Set ExcelSourceWorkbook = ExcelApplication.Application.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #Else
    Set ExcelSourceWorkbook = ExcelApplication.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #End If
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
    
    
    #If Mac Then
        DoEvents 'Mac needs a short delay on my machine
    #End If
    
    Set MailMergeSlide = ActivePresentation.Slides(MailMergeSlideNum).Duplicate
        
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

End Sub

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

Sub ManualMailMerge()

    Dim SlideShape  As Shape
    
    ProgressForm.Show
    
    ReDim ManualHeaders(0)
     
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes
            FindMergeFields SlideShape
        Next SlideShape
        
    Next PresentationSlide
   
    ProgressForm.Hide
    
    
    ManualHeaders = RemoveDuplicates(ManualHeaders)
    PreviewManualMailMerge.MailMergeListBox.Clear
    PreviewManualMailMerge.MailMergeListBox.ColumnCount = 2
    PreviewManualMailMerge.ReplaceTextTextBox.Text = ""
    PreviewManualMailMerge.ReplaceTextFrame.Caption = ""
    
    
    For HeaderCount = 0 To UBound(ManualHeaders) - 1
      PreviewManualMailMerge.MailMergeListBox.AddItem
      PreviewManualMailMerge.MailMergeListBox.List(HeaderCount, 0) = "{{" & ManualHeaders(HeaderCount) & "}}"
      PreviewManualMailMerge.MailMergeListBox.List(HeaderCount, 1) = ""
    Next HeaderCount
    
    CancelTriggered = False
    
    PreviewManualMailMerge.Show
    
    If CancelTriggered = True Then Exit Sub
    
    ReDim ManualTexts(UBound(ManualHeaders))
    
    For ManualTextCount = 0 To UBound(ManualHeaders) - 1
        ManualHeaders(ManualTextCount) = PreviewManualMailMerge.MailMergeListBox.List(ManualTextCount, 0)
        ManualTexts(ManualTextCount) = PreviewManualMailMerge.MailMergeListBox.List(ManualTextCount, 1)
    Next ManualTextCount
    
    ProgressForm.Show
       
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        
        For Each SlideShape In PresentationSlide.Shapes
            ReplaceMergeFields SlideShape, ManualHeaders, ManualTexts
        Next SlideShape
        
    Next PresentationSlide
   
    ProgressForm.Hide

End Sub

Sub FindMergeFields(SlideShape)
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            FindMergeFields SlideShapeChild
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            If Not SlideShape.TextFrame.TextRange = "" Then
                                   
                Set ShapeTextRange = SlideShape.TextFrame.TextRange

                  ''start
                
                  If ShapeTextRange.Words.Count > 2 Then
                      For WordCount = 2 To ShapeTextRange.Words.Count - 1
                      
                        If ShapeTextRange.Words(WordCount - 1) = "{{" And Left(ShapeTextRange.Words(WordCount + 1), 2) = "}}" Then
                        
                        If IsEmpty(ManualHeaders) Then
                            ReDim Preserve ManualHeaders(0)
                            ManualHeaders(0) = ShapeTextRange.Words(WordCount)
                        Else

                            ReDim Preserve ManualHeaders(UBound(ManualHeaders) + 1)
                            
                            ManualHeaders(UBound(ManualHeaders)) = ShapeTextRange.Words(WordCount)
                        End If
                        
                        End If
                      
                      Next WordCount
                      
                  End If
                  
                  ''stop
        
            End If
            
        End If
        
        If SlideShape.HasTable Then
            For TableRow = 1 To SlideShape.Table.Rows.Count
                For TableColumn = 1 To SlideShape.Table.Columns.Count
                    
                    If Not SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame.TextRange = "" Then
                            
                        Set ShapeTextRange = SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame.TextRange
                        
                      ''start
                    
                      If ShapeTextRange.Words.Count > 2 Then
                      For WordCount = 2 To ShapeTextRange.Words.Count - 1
                      
                        If ShapeTextRange.Words(WordCount - 1) = "{{" And Left(ShapeTextRange.Words(WordCount + 1), 2) = "}}" Then
                        
                        If IsEmpty(ManualHeaders) Then
                            ReDim Preserve ManualHeaders(0)
                            ManualHeaders(0) = ShapeTextRange.Words(WordCount)
                        Else

                            ReDim Preserve ManualHeaders(UBound(ManualHeaders) + 1)
                            
                            ManualHeaders(UBound(ManualHeaders)) = ShapeTextRange.Words(WordCount)
                        End If
                        
                        End If
                      
                      Next WordCount
                          
                      End If
                      
                      ''stop
                        
                    End If
                    
                Next
            Next
        End If
        
        If SlideShape.HasSmartArt Then
            
            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.Count
                
                For Each SlideSmartArtNode In SlideShape.SmartArt.AllNodes
                    
                    If Not SlideSmartArtNode.TextFrame2.TextRange = "" Then
                            
                        Set ShapeTextRange = SlideSmartArtNode.TextFrame2.TextRange

                      ''start
                    
                      If ShapeTextRange.Words.Count > 2 Then
                      For WordCount = 2 To ShapeTextRange.Words.Count - 1
                      
                        If ShapeTextRange.Words(WordCount - 1) = "{{" And Left(ShapeTextRange.Words(WordCount + 1), 2) = "}}" Then
                        
                        If IsEmpty(ManualHeaders) Then
                            ReDim Preserve ManualHeaders(0)
                            ManualHeaders(0) = ShapeTextRange.Words(WordCount)
                        Else

                            ReDim Preserve ManualHeaders(UBound(ManualHeaders) + 1)
                            
                            ManualHeaders(UBound(ManualHeaders)) = ShapeTextRange.Words(WordCount)
                        End If
                        
                        End If
                      
                      Next WordCount
                          
                      End If
                      
                      ''stop
   
                    End If
                    
                Next
                
            Next
            
        End If
        
    End If
    
End Sub

Sub ExcelFullFileMailMerge()
     
    Dim ExcelFile   As String
    Dim SlideShape  As Shape
    
    Dim ExcelApplication, ExcelSourceSheet, ExcelSourceWorkbook As Object
    
    Set ThisPresentation = ActivePresentation
    
    MailMergeSlideNum = ActiveWindow.Selection.SlideRange(1).SlideNumber
    
       
    #If Mac Then
    
    ExcelFile = MacFileDialog("/")
    
    If ExcelFile = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    
    #Else
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
    #End If
    
    On Error Resume Next
    Set ExcelApplication = GetObject(Class:="Excel.Application")
    Err.Clear
    If ExcelApplication Is Nothing Then Set ExcelApplication = CreateObject(Class:="Excel.Application")
    On Error GoTo 0
    
    
    #If Mac Then
    Set ExcelSourceWorkbook = ExcelApplication.Application.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #Else
    Set ExcelSourceWorkbook = ExcelApplication.Workbooks.Open(FileName:=ExcelFile, ReadOnly:=True)
    #End If
    Set ExcelSourceSheet = ExcelSourceWorkbook.Sheets(1)
    
    On Error GoTo HandleError
    Set LastCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=1, SearchDirection:=2).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, SearchOrder:=2, SearchDirection:=2).Column)
    Set FirstCell = ExcelSourceSheet.Cells(ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=1, SearchDirection:=1).Row, ExcelSourceSheet.Cells.Find(What:="*", LookIn:=-4163, After:=LastCell, SearchOrder:=2, SearchDirection:=1).Column)
    On Error GoTo 0
       
    Dim MergeFields() As Variant
    Dim MergeTexts() As Variant
    MergeFields = ExcelSourceSheet.Range(FirstCell.Address & ":" & ExcelSourceSheet.Cells(FirstCell.Row, LastCell.Column).Address).Value
    MergeTexts = ExcelSourceSheet.Range(ExcelSourceSheet.Cells(FirstCell.Row + 1, FirstCell.Column).Address & ":" & ExcelSourceSheet.Cells(LastCell.Row, LastCell.Column).Address).Value
    
    PreviewFullFileMailMerge.MailMergeHeadersListBox.Clear
    PreviewFullFileMailMerge.MailMergeHeadersListBox.ColumnCount = UBound(MergeFields, 2)
    PreviewFullFileMailMerge.MailMergeHeadersListBox.List = ExcelSourceSheet.Range(FirstCell.Address & ":" & ExcelSourceSheet.Cells(FirstCell.Row, LastCell.Column).Address).Value
    
    PreviewFullFileMailMerge.MailMergeListBox.Clear
    PreviewFullFileMailMerge.MailMergeListBox.ColumnCount = UBound(MergeFields, 2)
    PreviewFullFileMailMerge.MailMergeListBox.List = ExcelSourceSheet.Range(ExcelSourceSheet.Cells(FirstCell.Row + 1, FirstCell.Column).Address & ":" & LastCell.Address).Value
    
    PreviewFullFileMailMerge.ExampleLabel.Caption = "Data taken from the first sheet of the Excel-file. Current selected slide will be duplicated" & Str(UBound(MergeTexts, 1)) & " times and all mail merge fields placed between {{ }} will be replaced with the data above." & vbNewLine & vbNewLine & "Example: {{" & MergeFields(1, 1) & "}}" & " will be replaced with " & MergeTexts(1, 1) & " on the first slide."
    
    DotPosition = InStrRev(ActivePresentation.Name, ".")
        If DotPosition > 0 Then
            PresentationFilename = Left(ActivePresentation.Name, DotPosition - 1)
        Else
            PresentationFilename = ActivePresentation.Name
        End If
       
    PreviewFullFileMailMerge.MergeFilename.Text = PresentationFilename & " {{" & MergeFields(1, 1) & "}}"
    
    ExcelSourceWorkbook.Close
    
    CancelTriggered = False
    
    PreviewFullFileMailMerge.Show
    
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
    
    
    #If Mac Then
        DoEvents 'Mac needs a short delay on my machine
    #End If
    

    'Dim TempFilename As TextRange
    Dim SlidePlaceHolder As PowerPoint.Shape
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    SlidePlaceHolder.TextFrame.TextRange.Text = PreviewFullFileMailMerge.MergeFilename.Text
    Set TempFilename = SlidePlaceHolder.TextFrame.TextRange
    
                    For MergeFieldsCount = LBound(TempMergeFields) To UBound(TempMergeFields)
                    
                        Set TemporaryTextRange = TempFilename.Replace(FindWhat:=TempMergeFields(MergeFieldsCount), Replacewhat:=TempMergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                    
                    Do While Not TemporaryTextRange Is Nothing
                        Set TempFilename = TempFilename.Characters(TemporaryTextRange.Start + TemporaryTextRange.Length, TempFilename.Length)
                        Set TemporaryTextRange = TempFilename.Replace(FindWhat:=TempMergeFields(MergeFieldsCount), Replacewhat:=TempMergeTexts(MergeFieldsCount), WholeWords:=msoFalse)
                    Loop
                    
                Next MergeFieldsCount
                
    PresentationFilename = SlidePlaceHolder.TextFrame.TextRange.Text
    SlidePlaceHolder.Delete
    
    
    'Set MailMergeSlide = ActivePresentation.Slides(MailMergeSlideNum).Duplicate
        
    'For Each SlideShape In MailMergeSlide.Shapes
    '    ReplaceMergeFields SlideShape, TempMergeFields, TempMergeTexts
    'Next SlideShape
        
        #If Mac Then

        ThisPresentation.SaveCopyAs ActivePresentation.Path & "/" & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(ActivePresentation.Path & "/" & PresentationFilename & ".pptx")
        
        #Else
        
        ThisPresentation.SaveCopyAs ActivePresentation.Path & "\" & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(ActivePresentation.Path & "\" & PresentationFilename & ".pptx")
        
        #End If
        
        ProgressForm.Show
        
        NumberOfSlides = TemporaryPresentation.Slides.Count
        For SlideLoop = TemporaryPresentation.Slides.Count To 1 Step -1
            SetProgress ((NumberOfSlides - SlideLoop) / NumberOfSlides * 100)
            'code here
            
                For Each SlideShape In TemporaryPresentation.Slides(SlideLoop).Shapes
                  ReplaceMergeFields SlideShape, TempMergeFields, TempMergeTexts
                Next SlideShape
            
            
        Next SlideLoop
        
        ProgressForm.Hide
        
        TemporaryPresentation.Save
        TemporaryPresentation.Close
        
        
    Next i
    
    ProgressForm.Hide
    
    
    Exit Sub
    
HandleError:
    ExcelSourceWorkbook.Close
    MsgBox "Cannot load data. Does the first sheet in the Excel-file contain data with headers?"

End Sub
