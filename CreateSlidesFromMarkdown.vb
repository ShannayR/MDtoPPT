Sub CreateSlidesFromMarkdown()
    ' Variables for file handling
    Dim fd As FileDialog
    Dim fileName As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim useBulletSplitting As Boolean
    
    ' Ask user if they want bullet splitting
    Dim response As Integer
    response = MsgBox("Do you want to automatically split slides with more than 10 bullet points into multiple slides?", _
                     vbYesNo + vbQuestion, "Bullet Splitting Option")
    
    If response = vbYes Then
        useBulletSplitting = True
    Else
        useBulletSplitting = False
    End If
    
    ' Use PowerPoint's FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Markdown Files", "*.md"
        .Title = "Select Slideshow Outline File"
        If .Show = -1 Then
            fileName = .SelectedItems(1)
        Else
            MsgBox "No file selected."
            Exit Sub
        End If
    End With
    
    ' Read entire file
    fileNum = FreeFile
    Open fileName For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' Create presentation
    Dim ppt As Presentation
    Set ppt = Presentations.Add
    
    ' Use the Split method to break the file into slide sections
    Dim slideBlocks As Variant
    slideBlocks = Split(fileContent, "### Slide ")
    
    Dim slideCount As Integer
    slideCount = 0
    
    ' Process each slide block (start from 1 since element 0 is content before first slide)
    Dim i As Integer
    For i = 1 To UBound(slideBlocks)
        slideCount = slideCount + 1
        If useBulletSplitting Then
            Call ProcessSlideWithSplitting(ppt, CStr(slideBlocks(i)), slideCount)
        Else
            Call ProcessSlideSimple(ppt, CStr(slideBlocks(i)), slideCount)
        End If
    Next i
    
    ' Remove the default blank slide that PowerPoint creates
    If slideCount > 0 And ppt.Slides.Count > slideCount Then
        ppt.Slides(ppt.Slides.Count).Delete
    End If
    
    MsgBox "Successfully created " & slideCount & " slides!"
End Sub

' Simple processing without bullet splitting
Sub ProcessSlideSimple(ppt As Presentation, slideText As String, slideNum As Integer)
    Dim slide As slide
    Dim titleShape As shape
    Dim contentShape As shape
    
    ' Create slide
    Set slide = ppt.Slides.Add(slideNum, ppLayoutText)
    
    ' Split slide text into lines
    Dim lines As Variant
    slideText = Replace(slideText, vbCrLf, vbLf)
    slideText = Replace(slideText, vbCr, vbLf)
    lines = Split(slideText, vbLf)
    
    ' Extract title from first line
    Dim slideTitle As String
    Dim firstLine As String
    
    If UBound(lines) >= 0 Then
        firstLine = Trim(lines(0))
        If InStr(firstLine, ": ") > 0 Then
            slideTitle = Trim(Mid(firstLine, InStr(firstLine, ": ") + 2))
        Else
            slideTitle = "Slide " & slideNum
        End If
    Else
        slideTitle = "Slide " & slideNum
    End If
    
    ' Set title
    Set titleShape = slide.Shapes.Title
    titleShape.TextFrame.TextRange.Text = slideTitle
    
    ' Format title
    With titleShape.TextFrame.TextRange.Font
        .Name = "Calibri"
        .Size = 28
        .Bold = True
        .Color.RGB = RGB(68, 114, 196)
    End With
    
    ' Collect content
    Dim slideContent As String
    slideContent = ""
    Dim j As Integer
    
    If UBound(lines) > 0 Then
        For j = 1 To UBound(lines)
            Dim line As String
            line = Trim(lines(j))
            
            If Len(line) > 0 And line <> "---" And Left(line, 4) <> "### " Then
                If slideContent <> "" Then
                    slideContent = slideContent & vbCrLf
                End If
                line = Replace(line, "**", "")
                slideContent = slideContent & line
            End If
        Next j
    End If
    
    ' Set content
    Call SetSlideContent(slide, slideContent)
End Sub

' Processing with bullet splitting
Sub ProcessSlideWithSplitting(ppt As Presentation, slideText As String, ByRef slideNum As Integer)
    Dim lines As Variant
    slideText = Replace(slideText, vbCrLf, vbLf)
    slideText = Replace(slideText, vbCr, vbLf)
    lines = Split(slideText, vbLf)
    
    ' Extract title
    Dim slideTitle As String
    Dim firstLine As String
    
    If UBound(lines) >= 0 Then
        firstLine = Trim(lines(0))
        If InStr(firstLine, ": ") > 0 Then
            slideTitle = Trim(Mid(firstLine, InStr(firstLine, ": ") + 2))
        Else
            slideTitle = "Slide " & slideNum
        End If
    Else
        slideTitle = "Slide " & slideNum
    End If
    
    ' Separate bullets from other content
    Dim allBullets() As String
    Dim otherContent As String
    Dim bulletCount As Integer
    Dim j As Integer
    
    bulletCount = 0
    otherContent = ""
    ReDim allBullets(0)
    
    If UBound(lines) > 0 Then
        For j = 1 To UBound(lines)
            Dim line As String
            line = Trim(lines(j))
            
            If Len(line) > 0 And line <> "---" And Left(line, 4) <> "### " Then
                line = Replace(line, "**", "")
                
                ' Check if bullet point
                If Left(line, 2) = "- " Or Left(line, 2) = "* " Or Left(line, 2) = "+ " Or _
                   (Len(line) > 3 And Left(line, 3) >= "1. " And Left(line, 3) <= "9. ") Then
                    bulletCount = bulletCount + 1
                    ReDim Preserve allBullets(bulletCount - 1)
                    allBullets(bulletCount - 1) = line
                Else
                    If otherContent <> "" Then
                        otherContent = otherContent & vbCrLf
                    End If
                    otherContent = otherContent & line
                End If
            End If
        Next j
    End If
    
    ' Determine slides needed
    Const MAX_BULLETS_PER_SLIDE As Integer = 10
    Dim totalSlides As Integer
    
    If bulletCount <= MAX_BULLETS_PER_SLIDE Then
        totalSlides = 1
    Else
        totalSlides = Int((bulletCount - 1) / MAX_BULLETS_PER_SLIDE) + 1
    End If
    
    ' Create slides
    Dim currentSlideIndex As Integer
    For currentSlideIndex = 1 To totalSlides
        Dim slide As slide
        Set slide = ppt.Slides.Add(slideNum, ppLayoutText)
        
        ' Set title
        Dim currentSlideTitle As String
        If totalSlides > 1 Then
            currentSlideTitle = slideTitle & " - Part " & currentSlideIndex
        Else
            currentSlideTitle = slideTitle
        End If
        
        Dim titleShape As shape
        Set titleShape = slide.Shapes.Title
        titleShape.TextFrame.TextRange.Text = currentSlideTitle
        
        ' Format title
        With titleShape.TextFrame.TextRange.Font
            .Name = "Calibri"
            .Size = 28
            .Bold = True
            .Color.RGB = RGB(68, 114, 196)
        End With
        
        ' Build content
        Dim slideContent As String
        slideContent = ""
        
        ' Add other content to first slide
        If currentSlideIndex = 1 And Len(otherContent) > 0 Then
            slideContent = otherContent
        End If
        
        ' Add bullets for this slide
        If bulletCount > 0 Then
            Dim startBullet As Integer
            Dim endBullet As Integer
            
            startBullet = (currentSlideIndex - 1) * MAX_BULLETS_PER_SLIDE
            endBullet = startBullet + MAX_BULLETS_PER_SLIDE - 1
            If endBullet > bulletCount - 1 Then
                endBullet = bulletCount - 1
            End If
            
            Dim k As Integer
            For k = startBullet To endBullet
                If slideContent <> "" Then
                    slideContent = slideContent & vbCrLf
                End If
                slideContent = slideContent & allBullets(k)
            Next k
        End If
        
        ' Set content
        Call SetSlideContent(slide, slideContent)
        
        ' Increment slide number for next slide
        If currentSlideIndex < totalSlides Then
            slideNum = slideNum + 1
        End If
    Next currentSlideIndex
End Sub

' Helper function to set slide content
Sub SetSlideContent(slide As slide, slideContent As String)
    Dim contentShape As shape
    
    If slide.Shapes.Count >= 2 Then
        Dim i As Integer
        Set contentShape = Nothing
        For i = 1 To slide.Shapes.Count
            If slide.Shapes(i).Type = msoPlaceholder Then
                If slide.Shapes(i).PlaceholderFormat.Type = ppPlaceholderBody Or _
                   slide.Shapes(i).PlaceholderFormat.Type = ppPlaceholderObject Then
                    Set contentShape = slide.Shapes(i)
                    Exit For
                End If
            End If
        Next i
        
        If Not contentShape Is Nothing Then
            If Len(slideContent) > 0 Then
                contentShape.TextFrame.TextRange.Text = slideContent
                With contentShape.TextFrame.TextRange.Font
                    .Name = "Calibri"
                    .Size = 16
                End With
                contentShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.2
            Else
                contentShape.TextFrame.TextRange.Text = ""
            End If
        End If
    End If
End Sub


