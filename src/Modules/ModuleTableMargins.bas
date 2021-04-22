Attribute VB_Name = "ModuleTableMargins"
Sub TablesMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
End Sub

Sub TablesMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        .MarginBottom = .MarginBottom + 0.2
                        .MarginLeft = .MarginLeft + 0.2
                        .MarginRight = .MarginRight + 0.2
                        .MarginTop = .MarginTop + 0.2
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
End Sub

Sub TablesMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        If .MarginBottom >= 0.2 Then
                            .MarginBottom = .MarginBottom - 0.2
                        End If
                        If .MarginLeft >= 0.2 Then
                            .MarginLeft = .MarginLeft - 0.2
                        End If
                        If .MarginRight >= 0.2 Then
                            .MarginRight = .MarginRight - 0.2
                        End If
                        If .MarginTop >= 0.2 Then
                            .MarginTop = .MarginTop - 0.2
                        End If
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
End Sub
