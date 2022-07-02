Sub AllStocksAnalysis()

' make timer
Dim startTime As Single
Dim endTime  As Single


yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
       
       

 Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'header rows
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"


' array of tickers
Dim tickers(11) As String

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




Dim startingPrice As Single

Dim endingPrice As Single

Sheets(yearValue).Activate

'rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row


    ' loop through tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
   
       'loop through rows in data
            
        
        Sheets(yearValue).Activate
            
        For j = 2 To RowCount
            
            
        'totalVolume current ticker
        
        If Cells(j, 1).Value = ticker Then

        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
           
           
        'starting price for current ticker
           
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        startingPrice = Cells(j, 6).Value

        End If

        
        'ending price
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        endingPrice = Cells(j, 6).Value
        
        End If
       
     
     
     Next j
   
   
   Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   
   
  Next i
  
  
  
    'Formatting

    Worksheets("All Stocks Analysis").Activate


    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    
    
    
For i = 4 To 15

    If Cells(i, 3) > 0 Then

            'Change cell green
            Cells(i, 3).Interior.Color = vbGreen


    ElseIf Cells(i, 3) < 0 Then

            'Change cell red
            Cells(i, 3).Interior.Color = vbRed
            
    Else

            'Clear color
            Cells(i, 3).Interior.Color = xlNone

    End If

Next i




        endTime = Timer
        
        MsgBox "This code ran in " & (endTime - startTime) & " second for the year " & (yearValue)


End Sub


Sub ClearWorksheet()

    Cells.Clear

End Sub