Sub update_TOC()
Dim i, sectNumb As Long
Dim shp As Shape
 
For Each shp In Application.ActiveWindow.View.Slide.Shapes
    If shp.HasTable Then
        
        Do While ActivePresentation.SectionProperties.Count - 1 > shp.Table.Rows.Count
            shp.Table.Rows.Add
        Loop
    
        Do While ActivePresentation.SectionProperties.Count - 1 < shp.Table.Rows.Count
            shp.Table.Rows(1).Delete
        Loop
        
        sectNumb = ActivePresentation.SectionProperties.Count
        With ActivePresentation.SectionProperties
            For i = 2 To .Count
                shp.Table.Cell(i - 1, 1).Shape.TextFrame.TextRange.Text = .Name(i)
                shp.Table.Cell(i - 1, 2).Shape.TextFrame.TextRange.Text = .FirstSlide(i)
            Next i
        End With
        
        
        'add hyperlinks
        For i = 2 To sectNumb
        
            'section name col
            With shp.Table.Cell(i - 1, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink
                .SubAddress = shp.Table.Cell(i - 1, 2).Shape.TextFrame.TextRange.Text
                .TextToDisplay = shp.Table.Cell(i - 1, 1).Shape.TextFrame.TextRange.Text
            End With
          
            'slide number col
            With shp.Table.Cell(i - 1, 2).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink
                .SubAddress = shp.Table.Cell(i - 1, 2).Shape.TextFrame.TextRange.Text
                .TextToDisplay = shp.Table.Cell(i - 1, 2).Shape.TextFrame.TextRange.Text
            End With
        Next i
        
        Exit Sub
    End If
    
Next shp
 
End Sub

