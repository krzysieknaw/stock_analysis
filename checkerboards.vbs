Sub MakeCheckerboard()

Dim Row As Integer, Col As Integer
For Row = 1 To 8

If WorksheetFunction.IsOdd(Row) Then
For Col = 2 To 8 Step 2

Cells(Row, Col).Interior.Color = 255
Next Col

Else

For Col = 1 To 8 Step 2

Cells(Row, Col).Interior.Color = 255
Next Col
End If
Next Row
End Sub

