Sub AllStocksAnalysis()
    Dim yearValue As String
    
    yearValue = VBA.Interaction.InputBox("What year would you like to run the analysis on?")
    
    
   Worksheets("All Stocks Analysis").Activate
   
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Dim tickers(12) As String
    
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'initialize variables and get row count
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    'make sure data worksheet activates
    Worksheets(yearValue).Activate
    
    'establish number of rows
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (RowCount)

    'need tickers from worksheet
    'ticker loop/outer for loop
    For i = 0 To 11
        ticker = tickers(i)
        'totalVolume = 0
        
        'reactivate data sheet
        Worksheets(yearValue).Activate
            'inner for
            For j = 2 To RowCount


                
            'Total Volume
            If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
            
            'Starting Price
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                    
            'Ending Price
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
            Next j
            
        'output
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
     
    'Formatting text
  
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit

    
   
End Sub