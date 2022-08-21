# stocks-analysis
Doing a stocks analysis using VBA.
## Overview of Project
The purpose of the project was to refactor a VBA code for a stock anlysis from the the years 2017 and 2018. This process was completed within module 2 already, but refactoring the code made it more efficient. The data had twelve stocks that contained ticker values, the date it was issued, the opening, closing, adjusted closing price, the highest and lowest price, and volume of the stock. The goal was to obtain the total daiily volume and the return percentage of each stock.
## Results
I started with adding the refactoring instructions to my code and then followed the steps to create the following code. Also, I got the code that was needed for the input box, chart headers, and ticker array and I copied those in the correct place to finish the whole code.
```
'1a) Create a ticker Index
    tickerIndex = 0

'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i

''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
 ```
When I ran the code the time was much faster than the original code that I created in the green stocks analysis.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/110861876/185811141-11dc5318-c8c2-4b0c-afa1-d2347a876439.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/110861876/185811147-ed88ce04-3002-4b57-92d6-c9a1b788a63b.png)
## Results
When refactoring code it makes it more organized and easier to read. It also runs faster to run but there are disadvantages like having too large of an application or not having the proper test cases for the existing codes. With the original code it takes longer for it to run and it seems like there is more information than needed once you do a refactored code, but the original is good for teaching introductory coders.
