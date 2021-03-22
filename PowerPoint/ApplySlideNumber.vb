Sub ApplySlideNumber()
    Dim oSl As Slide
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.Shapes
            If IsRegexMatch(oSh.Name, "Slide\sNumber\sPlaceholder\s\d+") Or oSh.Name = "SlideNumber" Then
              oSh.TextFrame.TextRange.Text = CStr(oSl.SlideIndex) & "/" _
                      & CStr(ActivePresentation.Slides.Count)
            End If
        Next
    Next
End Sub
