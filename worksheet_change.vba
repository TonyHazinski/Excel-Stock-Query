Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 Then
        Dim sheet As Worksheet
        Set sheet = Worksheets("Sheet1")
        If Not IsEmpty(sheet.Cells(Target.Row, 1)) Then
            updateStock (Target.Row)
        End If
    End If
    Target.Font.ColorIndex = 5
End Sub

