Sub persianSlideNumber()
    ActivePresentation.Slides(1).Shapes("slideNumberPe").Copy
    a = ActivePresentation.Slides(1).Shapes("slideNumberPe").TextFrame.TextRange.Text
    Dim i As Long
    
    For i = 2 To ActivePresentation.Slides.Count
        On Error Resume Next
        ActivePresentation.Slides(i).Shapes("slideNumberPe").Delete
        ActivePresentation.Slides(i).Shapes.Paste
        ActivePresentation.Slides(i).Shapes("slideNumberPe").TextFrame.TextRange.Text = a + (i - 1)
    Next
End Sub
