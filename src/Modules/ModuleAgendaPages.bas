Attribute VB_Name = "ModuleAgendaPages"
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

Sub CreateOrUpdateMasterAgenda()
    Dim NumberOfSections As Long
    Set MyDocument = Application.ActiveWindow
    Dim MasterExists As Boolean
    hasMasterAgenda = False
    Dim AgendaSlide As Slide
    Dim AgendaLayout As CustomLayout
    Dim AgendaShape As shape
    
    If ActivePresentation.SectionProperties.Count >= 2 Then
        
        'Check if master slide already exists
        For SlideLoop = ActivePresentation.Slides.Count To 1 Step -1
            If ActivePresentation.Slides(SlideLoop).Tags("INSTRUMENTA MASTER AGENDA PAGE") = "YES" Then
                
                hasMasterAgenda = True
                
                Set AgendaSlide = ActivePresentation.Slides(SlideLoop)
                
                For ShapeLoop = 1 To AgendaSlide.Shapes.Count
                    
                    If AgendaSlide.Shapes(ShapeLoop).Tags("INSTRUMENTA AGENDA TEXTSHAPE") = "YES" Then
                        
                        Set AgendaShape = AgendaSlide.Shapes(ShapeLoop)
                        Set OldAgendaShape = AgendaShape.Duplicate
                        
                        With OldAgendaShape
                            .left = Application.ActivePresentation.PageSetup.SlideWidth + 10
                        End With
                        
                    End If
                    
                Next ShapeLoop
                
            End If
            
        Next SlideLoop
        
        'If master does not exist, create one
        If hasMasterAgenda = False Then
            
            If ActivePresentation.Slides.Count = 0 Then
                Set AgendaSlide = ActivePresentation.Slides.AddSlide(1, ActivePresentation.SlideMaster.CustomLayouts(1))
            ElseIf ActivePresentation.Slides.Count = 1 Then
                Set AgendaSlide = ActivePresentation.Slides.AddSlide(2, ActivePresentation.Slides(1).CustomLayout)
            Else
                Set AgendaSlide = ActivePresentation.Slides.AddSlide(2, ActivePresentation.Slides(2).CustomLayout)
            End If
            
            Set AgendaShape = AgendaSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, Application.ActivePresentation.PageSetup.SlideWidth - 200, 50)
            AgendaShape.Name = "AgendaTextBox"
            AgendaSlide.Tags.Add "INSTRUMENTA MASTER AGENDA PAGE", "YES"
            AgendaShape.Tags.Add "INSTRUMENTA AGENDA TEXTSHAPE", "YES"
        End If
        
        With ActivePresentation.SectionProperties
            
            For NumberOfSections = 2 To .Count
                
                If NumberOfSections = 2 Then
                    AgendaShape.TextFrame.TextRange.Text = .Name(NumberOfSections)
                Else
                    AgendaShape.TextFrame.TextRange.Text = AgendaShape.TextFrame.TextRange.Text & vbNewLine & .Name(NumberOfSections)
                End If
                
            Next
            With AgendaShape.TextFrame.TextRange
                
                If hasMasterAgenda = False Then
                    With .Font
                        .Color.RGB = RGB(0, 0, 0)
                        .Bold = msoFalse
                        .Size = 16
                        .Italic = msoFalse
                        .Underline = msoFalse
                        .Emboss = msoFalse
                        .Name = "Arial"
                    End With
                    
                    With .Lines(1).Font
                        .Color.RGB = RGB(0, 51, 153)
                        .Bold = msoTrue
                        .Size = 16
                        .Italic = msoFalse
                        .Underline = msoFalse
                        .Emboss = msoFalse
                        .Name = "Arial"
                    End With
                    
                Else
                    With .Font
                        .Color.RGB = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Color.RGB
                        .Bold = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Bold
                        .Size = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Size
                        .Italic = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Italic
                        .Underline = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Underline
                        .Emboss = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Emboss
                        .Name = OldAgendaShape.TextFrame.TextRange.Lines(2).Font.Name
                    End With
                    
                    With .Lines(1).Font
                        .Color.RGB = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Color.RGB
                        .Bold = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Bold
                        .Size = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Size
                        .Italic = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Italic
                        .Underline = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Underline
                        .Emboss = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Emboss
                        .Name = OldAgendaShape.TextFrame.TextRange.Lines(1).Font.Name
                    End With
                End If
                
            End With
            
        End With
        
        AgendaSlide.MoveToSectionStart (2)
        
        If hasMasterAgenda = False Then
            CreateAgendaPages
            MsgBox "Agenda pages created." & vbNewLine & vbNewLine & "You can customize by updating the first agenda page (section 2) And then run this command again." & vbNewLine & vbNewLine & "All agenda pages will be formatted according to that first agenda page. The formatting of the first line (e.g. color, bold, italic) will be used to highlight the agenda-item on all other agenda pages."
        Else
            OldAgendaShape.Delete
            CreateAgendaPages
            MsgBox "Agenda slides updated."
        End If
        
    Else
        
        Dim HelpRequired As Integer
        Dim DoneWithCreatingSections As Integer
        Dim SectionToCreate As String
        Dim SectionsReady As Boolean
        Dim SectionNum As Long
        SectionsReady = False
              
        If ActivePresentation.SectionProperties.Count = 0 Then
            ActivePresentation.SectionProperties.AddSection 1
        End If
        
        SectionNum = 1
        
        HelpRequired = MsgBox("Your document only has one section. Please create a section for every agenda item you want to create and then run this command again. Note: The first section will be skipped." & vbNewLine & vbNewLine & "Do you want me to help you to create menu-items/sections?", vbQuestion + vbYesNo + vbDefaultButton1, "No sections found")
        
        If HelpRequired = vbYes Then
            
            Do While SectionsReady = False
                SectionToCreate = InputBox("Enter title for menu-item / section " & Str(SectionNum) & vbNewLine & vbNewLine & "Note: Use one or more '-' directly in front of the title to create different levels of subitems" & vbNewLine & vbNewLine & "Cancel or close this dialog when you're done.", "Enter title of menu-item / section")
                
                If StrPtr(SectionToCreate) = 0 Then
                    
                    SectionsReady = True
                    
                ElseIf SectionToCreate = vbNullString Then
                    SectionsReady = True
                    
                Else
                    SectionNum = SectionNum + 1
                    ActivePresentation.SectionProperties.AddSection SectionNum, SectionToCreate
                End If
                
            Loop
            
            MsgBox "Move your slides to the appropriate sections and run this command again to generate the agenda pages."
            
        End If
        
    End If
    
End Sub

Sub CreateAgendaPages()
    Dim NumberOfSections As Long
    Set MyDocument = Application.ActiveWindow
    
    Dim hasMasterAgenda As Boolean
    hasMasterAgenda = False
    Dim MasterAgendaSlide As Slide
    
    Dim AgendaTextBox As shape
    
    For SlideLoop = ActivePresentation.Slides.Count To 1 Step -1
        
        If ActivePresentation.Slides(SlideLoop).Tags("INSTRUMENTA AGENDA PAGE") = "YES" Then
            
            ActivePresentation.Slides(SlideLoop).Delete
            
        End If
        
    Next SlideLoop
    
    For SlideLoop = 1 To ActivePresentation.Slides.Count
        
        If ActivePresentation.Slides(SlideLoop).Tags("INSTRUMENTA MASTER AGENDA PAGE") = "YES" Then
            
            For ShapeLoop = 1 To ActivePresentation.Slides(SlideLoop).Shapes.Count
                
                If ActivePresentation.Slides(SlideLoop).Shapes(ShapeLoop).Tags("INSTRUMENTA AGENDA TEXTSHAPE") = "YES" Then
                    
                    Set MasterAgendaSlide = ActivePresentation.Slides(SlideLoop)
                    
                    Set MasterAgendaTextBox = ActivePresentation.Slides(SlideLoop).Shapes(ShapeLoop)
                    hasMasterAgenda = True
                    
                End If
                
            Next ShapeLoop
            
        End If
    Next SlideLoop
    
    If hasMasterAgenda = True Then
    
        
        For IndentLoop = 1 To MasterAgendaTextBox.TextFrame2.TextRange.Lines.Count
        MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).ParagraphFormat.IndentLevel = 1
        Next IndentLoop
        
        For IndentLoop = 1 To MasterAgendaTextBox.TextFrame2.TextRange.Lines.Count
        
        For DepthLoop = 1 To 6
        
        If MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).Characters(1, 1) = "-" Then
        MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).ParagraphFormat.Bullet.Type = msoBulletUnnumbered
        MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).ParagraphFormat.IndentLevel = MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).ParagraphFormat.IndentLevel + 1
        MasterAgendaTextBox.TextFrame2.TextRange.Lines(IndentLoop).Characters(1, 1).Delete
        End If
        
        Next DepthLoop
        
        Next IndentLoop
        
        For NumberOfSections = 2 To ActivePresentation.SectionProperties.Count - 1
            Set NewSlide = MasterAgendaSlide.Duplicate
            NewSlide.Tags.Add "INSTRUMENTA MASTER AGENDA PAGE", "NO"
            NewSlide.Tags.Add "INSTRUMENTA AGENDA PAGE", "YES"
            NewSlide.MoveToSectionStart (NumberOfSections + 1)
            
            For ShapeLoop = 1 To NewSlide.Shapes.Count
                
                If NewSlide.Shapes(ShapeLoop).Tags("INSTRUMENTA AGENDA TEXTSHAPE") = "YES" Then
                    Set AgendaTextBox = NewSlide.Shapes(ShapeLoop)
                End If
                
            Next ShapeLoop
            
            With AgendaTextBox.TextFrame.TextRange.Lines(1).Font
                .Color.RGB = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Color.RGB
                .Bold = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Bold
                .Size = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Size
                .Italic = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Italic
                .Underline = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Underline
                .Emboss = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Emboss
                .Name = MasterAgendaTextBox.TextFrame.TextRange.Lines(2).Font.Name
            End With
            
            With AgendaTextBox.TextFrame.TextRange.Lines(NumberOfSections).Font
                .Color.RGB = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Color.RGB
                .Bold = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Bold
                .Size = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Size
                .Italic = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Italic
                .Underline = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Underline
                .Emboss = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Emboss
                .Name = MasterAgendaTextBox.TextFrame.TextRange.Lines(1).Font.Name
            End With
            
        Next NumberOfSections
        
    Else
        MsgBox "This document has no master agenda slide. Please create one first."
    End If
    
End Sub
