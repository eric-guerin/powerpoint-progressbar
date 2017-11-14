Sub AutoSections()

     
    Dim intSlide As Integer
    Dim strFileName As String
    Dim strTemp As String
    Dim strNotes As String
    Dim tabSectionNames() As String
    Dim tabSectionSlides() As Integer
    Dim sec As Integer
    Dim secNumber As Integer
    sec = 1
    Dim thisSection As Integer
    Dim texte As String
    Dim debut As Integer
    Dim longueur As Integer
   
    Dim width As Integer
    Dim barWidth As Integer
    Dim Sep As String
    Dim normalColor
    Dim emphColor
    Dim font As String
    Dim size As Integer
    Dim j As Integer
    Dim visibleSlides As Integer

    visibleSlides = 0

    'parameters are here

    Sep = " " '" " to put the menu in a single line, vbCr to put the menu in vertical mode
    normalColor = RGB(173, 185, 202)
    emphColor = RGB(68, 84, 106)
    BackgroundColor = RGB(222, 235, 247)
    size = 14
    font = "Eurostile"

    width = ActivePresentation.PageSetup.SlideWidth

    With ActivePresentation
      
        For intSlide = 1 To .Slides.Count
            strNotes = ActivePresentation.Slides(intSlide).NotesPage. _
                Shapes.Placeholders(2).TextFrame.TextRange.Lines(1).Text
            strNotes = Replace(strNotes, vbLf, "")
            strNotes = Replace(strNotes, vbCr, "")

            If InStr(strNotes, "Section:") = 1 Then
                ReDim Preserve tabSectionNames(sec + 1)
                ReDim Preserve tabSectionSlides(sec + 1)
                tabSectionNames(sec) = Mid(strNotes, 9)
                tabSectionSlides(sec) = intSlide
                sec = sec + 1
            End If
            If Not ActivePresentation.Slides(intSlide).SlideShowTransition.Hidden = True Then
                visibleSlides = visibleSlides + 1
            End If
        Next intSlide
        secNumber = sec - 1
        For intSlide = 1 To .Slides.Count
            texte = ""
            For sec = 1 To secNumber
                thisSection = 0
                If intSlide >= tabSectionSlides(sec) Then
                    If sec = secNumber Then
                        thisSection = 1
                    ElseIf intSlide < tabSectionSlides(sec + 1) Then
                        thisSection = 1
                    End If
                End If
                If thisSection = 1 Then
                    debut = Len(texte) + 1
                    longueur = Len(tabSectionNames(sec)) + 1
                End If
                If Not texte = "" Then
                    texte = texte & Sep
                End If
                texte = texte & tabSectionNames(sec)
            Next sec
            j = 1
            While j <= ActivePresentation.Slides(intSlide).Shapes.Count
                If ActivePresentation.Slides(intSlide).Shapes(j).Name = "MyText" Then
                    ActivePresentation.Slides(intSlide).Shapes(j).Delete
                
                ElseIf ActivePresentation.Slides(intSlide).Shapes(j).Name = "MyBar1" Then
                    ActivePresentation.Slides(intSlide).Shapes(j).Delete
                
                ElseIf ActivePresentation.Slides(intSlide).Shapes(j).Name = "MyBar2" Then
                    ActivePresentation.Slides(intSlide).Shapes(j).Delete
                Else
                    j = j + 1
                End If
            Wend
            If intSlide >= tabSectionSlides(1) Then
                With ActivePresentation.Slides(intSlide).Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                        Left:=30, Top:=16, width:=width - 60, Height:=50)
                    With .TextFrame.TextRange
                        .Text = texte
                        .font.size = size
                        .font.Name = font
                        .font.Color.RGB = normalColor
                        .Characters(debut, longueur).font.Bold = True
                        .Characters(debut, longueur).font.Color.RGB = emphColor
                    End With
                    .Name = "MyText"
                End With
                With ActivePresentation.Slides(intSlide).Shapes.AddShape(Type:=msoShapeRectangle, _
                        Left:=30, Top:=38, width:=width - 60, Height:=3)
                    .Fill.BackColor.RGB = BackgroundColor
                    .Fill.ForeColor.RGB = BackgroundColor
                    .Line.Visible = False
                    .Name = "MyBar1"
                End With
                barWidth = CInt((width - 62#) * (intSlide - 1#) / (visibleSlides - 1#))
                With ActivePresentation.Slides(intSlide).Shapes.AddShape(Type:=msoShapeRectangle, _
                        Left:=31, Top:=38, width:=barWidth, Height:=3)
                    .Fill.BackColor.RGB = emphColor
                    .Fill.ForeColor.RGB = emphColor
                    .Line.Visible = False
                    .Name = "MyBar2"
                End With
            
            End If

        Next intSlide

    End With


End Sub
