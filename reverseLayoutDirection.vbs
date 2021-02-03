rem A method for reversing layout direction in Microsoft Office Powerpoint
rem Usecases: Change LTR theme to RTL theme, Change RTL theme to LTR theme

Sub reverseLayout()
    ActiveWindow.Selection.ShapeRange.Flip (msoFlipHorizontal)
    SlideWidth = Application.ActivePresentation.PageSetup.SlideWidth
    For Each Shp In ActiveWindow.Selection.ShapeRange
        ShapeLeft = Shp.Left
        ShapeWidth = Shp.Width
        Shp.Left = SlideWidth - (ShapeLeft + ShapeWidth)
    Next
End Sub
