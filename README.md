# Overview of Project
### Our customer Steve is doing financial research for his parents on a green energy company DAQO. Steve wants to study both DAQO and some comparable stock tickers to encourage his parents to diversify.

### The ask from Steve is to create a VBA macro that he can use to compare stock data. Steve would also like to see the speed at which the code runs so that he can measure code performance on different datasets.

# Results
## Tables
### The tables below show the results of our macro for each year being studied. This should give Steve plenty of information to help his parents. 
![table of results 2017](https://github.com/marveld21/stocks-analysis/blob/main/Reources/stock_table_2017.png "Results for 2017")
![table of results 2018](https://github.com/marveld21/stocks-analysis/blob/main/Reources/stock_table_2018.png "Results for 2018")

## Code Refactoring
### The original code when run with the 2017 and 2018 data completed in 0.644 seconds and 0.640 seconds respectively. Refactoring the code using arrays to assist increased efficiency and resulted in the code being run a little more than 8 times faster. Refactored code ran for 2017 and 2018 for 0.078 seconds and 0.074 seconds respectively.

#### Original code overwriting with each loop
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   Next i

#### Refactored code using arrays to store values and writing once when finished
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        

        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        'End If

        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
        'If  Then
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
               
            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        'End If
        
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1


# Summary of Code Refactoring
### Refactoring code increases code runtime efficiency and makes code easier to read but it can make it more difficult to write initially. A good strategy is to get a piece of code working, then return to the code to refactor.

### Our original script worked correctly and did not take very long to run. Our refactored script took much longer to write but runs much faster. With such a small dataset it probably was unnecessary but as the data scales up the refactored code will increase in value.
