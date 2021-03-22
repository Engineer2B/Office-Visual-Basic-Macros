Sub ApplyFooter()
    Dim oSl As Slide
    Dim strTitle As String
    Dim strSubTitle As String
    strSubTitle = ""
    Dim strDate As String
    strDate = ""
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.Shapes
            If IsRegexMatch(oSh.Name, "(Titel|Title)\s\d+") Then
                strTitle = oSh.TextFrame.TextRange.Text
            End If
            If IsRegexMatch(oSh.Name, "(Ondertitel|Subtitle)\s\d+") Then
                strSubTitle = oSh.TextFrame.TextRange.Text
            End If
            If IsRegexMatch(oSh.Name, "(Datum|Date)") Then
                strDate = oSh.TextFrame.TextRange.Text
            End If
        Next
        If strTitle <> "" Then Exit For
    Next
    If strSubTitle <> "" Then
        strTitle = strTitle & " - " & strSubTitle
    End If
    If strDate <> "" Then
        strTitle = strTitle & " " & strDate
    End If

    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.Shapes
            If IsRegexMatch(oSh.Name, "Footer\sPlaceholder\s\d+") Then
                oSh.TextFrame.TextRange.Text = strTitle
                Exit For
            End If
        Next
    Next
End Sub
