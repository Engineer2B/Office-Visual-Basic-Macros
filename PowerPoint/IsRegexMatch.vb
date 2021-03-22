Function IsRegexMatch(ByVal vStr As String, ByVal patternStr As String)
    Dim myRegex As New RegExp
    With myRegex
        .Global = True
        .Multiline = True
        .IgnoreCase = False
        .pattern = patternStr
    End With
    IsRegexMatch = myRegex.Test(vStr)
End Function

