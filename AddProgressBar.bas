Attribute VB_Name = "AddProgressBar"
Sub AddProgressBar()
    On Error Resume Next
    With ActivePresentation
        For X = 2 To .Slides.Count
            ' Delete the existing progress bar background, if it exists
            .Slides(X).Shapes("PBBG").Delete
            
            ' Add a new shape as the progress bar background
            Set s = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
                0, .PageSetup.SlideHeight - 5, _
                .PageSetup.SlideWidth, 5)
            s.Fill.ForeColor.RGB = RGB(82, 197, 235)
            s.Line.Visible = False
            s.Name = "PBBG"
            
            
            ' Delete the existing progress bar, if it exists
            .Slides(X).Shapes("PB").Delete
            
            ' Add a new shape as the progress bar
            Set s = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
                0, .PageSetup.SlideHeight - 5, _
                (X - 1) * .PageSetup.SlideWidth / (.Slides.Count - 1), 5)
            s.Fill.ForeColor.RGB = RGB(46, 131, 195)
            s.Line.Visible = False
            s.Name = "PB"
        Next X:
    End With
End Sub


