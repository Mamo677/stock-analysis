# stock-analysis
##Over view
The purpose of this project was to make an efficient way to look at multiple stocks using VBA. 
After being able to run the analysis of these 12 different stocks the first, 
it was apparent that there was a more efficient way to work with the data given. In order to make the analysis more efficient, 
I needed to refactor my code. This project looks to see if my refactoring made the analysis more efficient.

## result
 # Refactored code  AllStocksAnalysisRefactored
Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop  over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
     For i = 0 To 11
    tickerIndex = tickers(1)

    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As String
    Dim tickerEndingPrice  As String
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    Worksheets(yearValue).Activate
         tickerVolumes = 0

         
         'If Cells(J,1).Value=tickerIndex
         'Then
    ''2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount
    'If Cells(j,1).value=tickerIndex
    
        '3a) Increase volume for current ticker
        
         If Cells(j, 1).Value = tickerIndex Then
      
        tickerVolumes = tickerVolumes + Cells(j, 8).Value

       End If

        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           tickerStartingPrices = Cells(j, 6).Value
           
        End If


            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           tickerEndingPrices = Cells(j, 6).Value

       End If

            

            '3d Increase the tickerIndex.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           tickerStartingPrices = Cells(j, 6).Value
           
        End If

            
        'End If
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value Then
        tickerEndingPrice = Cells(j, 6).Value
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate

        Cells(4, i, 1).Value = tickerIndex
        Cells(4, i, 2).Value = tickerVolumes
        Cells(4, i, 3).Value = tickerEndingPrice / tickerStartingPrices
        
        
    Next i
    
    'Formatting
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
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


## summery
1.The major advantage of refactoring code in VBA script is that you can use as much as of the original code as you want to and can put your new code side by side with your old code using different modules. The major disadvantage of refactoring code in VBA script is that if you do not have a strong understanding of the syntax, you will struggle to refactor your code as the syntax matters so much more when trying to make your code more efficient.

2. The major advantage of refactoring code is making the code more efficient. The major disadvantage of refactoring code is that you are taking code that already works and potential making it unusable if you can refactor it correctly. For that reason it is always smart to save your original code just incase you end up not being able to refactor it.
