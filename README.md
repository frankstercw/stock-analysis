# stock-analysis

Sub DQAnalysis()

yearValue = InputBox(2017)

   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Year"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

    Dim tickerIndex As String

    tickerIndex = 0


   '2) Initialize array of all tickers
   Dim tickers(12) As String
   tickers(tickerIndex) = "AY"
   tickers(tickerIndex + 1) = "CSIQ"
   tickers(tickerIndex + 2) = "DQ"
   tickers(tickerIndex + 3) = "ENPH"
   tickers(tickerIndex + 4) = "FSLR"
   tickers(tickerIndex + 5) = "HASI"
   tickers(tickerIndex + 6) = "JKS"
   tickers(tickerIndex + 7) = "RUN"
   tickers(tickerIndex + 8) = "SEDG"
   tickers(tickerIndex + 9) = "SPWR"
   tickers(tickerIndex + 10) = "TERP"
   tickers(tickerIndex + 11) = "VSLR"
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers

   For i = tickerIndex To 11
       tickerIndex = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = tickerIndex Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickerIndex
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

End Sub

Sub Format()

Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

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

End Sub
