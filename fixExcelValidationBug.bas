Private Sub Worksheet_Change(ByVal Target As Range)
    For r = 1 To Target.Rows.Count
        For c = 1 To Target.Columns.Count
            If Target.Cells(r, c) <> "" And Not Target.Cells(r, c).Validation.Value Then
                Target.Cells(r, c) = ""
            End If
        Next c
    Next r
End Sub
