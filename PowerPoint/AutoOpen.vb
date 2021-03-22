Sub Auto_Open()
'Note this will only fire automatically when in a loaded Add-in.
'
    Dim strPath As String
    For Each ad In AddIns
        If ad.Name = "Macros-add-in" Then
          strPath = ad.Path
        End If

    Next

     'Path for file with macros:
    Dim strFileName As String
    strFileName = strPath & "\Macros.pptm"

     'Open the presentation with macros, but keep it hidden:
    Application.Presentations.Open FileName:=strFileName, WithWindow:=msoFalse
End Sub
