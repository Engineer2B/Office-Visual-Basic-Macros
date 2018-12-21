Sub PrintNotes()
  Dim oSlides As Slides
  Dim oSl As Slide
  Dim oSh As Shape
  Dim strNotesText As String
  Dim strFileName As String
  Dim intFileNum As Integer
  Dim lngReturn As Long
  
  ' Get a filename to store the collected text
  strFileName = InputBox("Enter the full path and name of file to extract notes text to", "Output file?")
  
  exportHTML = True
  exportTXT = True
  
  ' did user cancel?
  If strFileName = "" Then
      Exit Sub
  End If
  
  If strFileName = "default" Or strFileName = "cwd" Then
    strFileName = ActivePresentation.Path & "\notes"
  End If
  
  ' is the path valid?  crude but effective test:  try to create the file.
  If exportHTML Or exportTXT Then
    If exportHTML Then
      fileExt = ".html"
    Else
      fileExt = ".txt"
    End If

    intFileNum = FreeFile()
    On Error Resume Next
    Open strFileName & fileExt For Output As intFileNum
    If Err.Number <> 0 Then     ' we have a problem
        MsgBox "Couldn't create the file: " & strFileName & vbCrLf _
            & "Please try again."
        Exit Sub
    End If
    Close #intFileNum  ' temporarily
  End If

  ' Get the notes text
  Set oSlides = ActivePresentation.Slides
  If exportHTML = True Then
    For Each oSl In oSlides
        For Each oSh In oSl.NotesPage.Shapes
        If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
            If oSh.HasTextFrame Then
                If oSh.TextFrame.HasText Then
                  strNotesTextHTML = strNotesTextHTML & SlideAsHTML(oSl, oSh.TextFrame)
                End If
            End If
        End If
        Next oSh
    Next oSl
    ' now write the text to file
    strOutFileName = strFileName & ".html"
    Open strOutFileName For Output As intFileNum
    Print #intFileNum, strNotesTextHTML
    Close #intFileNum
    
    ' show what we've done
    lngReturn = Shell("explorer " & strOutFileName, vbNormalFocus)
  End If
  
  If exportTXT = True Then
    For Each oSl In oSlides
        For Each oSh In oSl.NotesPage.Shapes
        If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
            If oSh.HasTextFrame Then
                If oSh.TextFrame.HasText Then
                  strNotesText = strNotesText & SlideTitle(oSl) & vbCrLf _
                    & TextFromLines(oSh.TextFrame) & vbCrLf & vbCrLf
                End If
            End If
        End If
        Next oSh
    Next oSl
    ' now write the text to file
    strOutFileName = strFileName & ".txt"
    Open strOutFileName For Output As intFileNum
    Print #intFileNum, strNotesText
    Close #intFileNum
    
    ' show what we've done
    lngReturn = Shell("explorer " & strOutFileName, vbNormalFocus)
  End If
End Sub


Function SlideAsHTML(ByRef inSlide As Slide, ByRef inTextFrame As TextFrame) As String
   SlideAsHTML = "<h1>" & SlideTitle(inSlide) & "</h1>" & vbCrLf _
   & "<p>" & TextFromLines(inTextFrame) & "</p>" & vbCrLf & vbCrLf
End Function

Function SlideAsPlainText(ByRef inSlide As Slide, ByRef inTextFrame As TextFrame) As String
   SlideAsPlainText = SlideTitle(oSl) & vbCrLf _
   & TextFromLines(oSh.TextFrame) & vbCrLf & vbCrLf
End Function


Function TextFromLines(ByRef inTextFrame As TextFrame) As String
  For Each oLine In inTextFrame.TextRange.Lines
    TextFromLines = TextFromLines & oLine.Text & vbCrLf
  Next oLine
End Function

Function SlideTitle(oSl As Slide) As String
  Dim oSh As Shape
  For Each oSh In oSl.Shapes
    If oSh.Type = msoPlaceholder Then
      If oSh.PlaceholderFormat.Type = ppPlaceholderTitle _
        Or oSh.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
        If Len(oSh.TextFrame.TextRange.Text) > 0 Then
            SlideTitle = "Slide " & CStr(oSl.SlideIndex) & ": " & oSh.TextFrame.TextRange.Text
        Else
            SlideTitle = "Slide " & CStr(oSl.SlideIndex)
        End If
        Exit Function
      End If
    End If
  Next
  SlideTitle = "Slide " & CStr(oSl.SlideIndex)
End Function
