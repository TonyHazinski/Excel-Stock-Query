Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 Then
        Dim sheet As Worksheet
        Set sheet = Worksheets("Sheet1")
        updateStock (Target.Row)
'        sheet.Cells(20, 4).Formula = sheet.Cells(20, 4).Formula
    End If
    Target.Font.ColorIndex = 5
End Sub
