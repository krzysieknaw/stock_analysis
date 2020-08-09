    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingPrice As Single
    Dim endingPrice As Single

    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        Worksheets(yearValue).Activate

        'loop over all the rows
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then

                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value

            End If

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If

        Next j