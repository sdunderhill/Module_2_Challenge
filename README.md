# VBA of Wall Street
## Overview

I originally received data from Steve. He was helping his parents look over the stock data for DAQO, a green energy company they wanted to invest all their money in. Through the process I also helped him to be able to analyze all the data on the spreadsheet for each given year. The end result is to show the return of the given stocks for 2017 and 2018. I created a script that ran that analysis but discovered there was a different way to write the script so it would run faster.

### Results of Year Comparison

The stocks all faired better in 2017 than in 2018. All of the stocks had a major decline in return in 2018 compared to 2017.  Only 2 stocks showed a positive return in 2018. Neither of which were the stock that Steve's parents chose. 

2017 had only one stock with a negative return.

![2017_returns](https://user-images.githubusercontent.com/81715217/118377866-87523b80-b595-11eb-8e9a-c44f989e7115.png)

2018 only had two positive returns

![2018_returns](https://user-images.githubusercontent.com/81715217/118377881-a224b000-b595-11eb-8e3f-9c16deeed16a.png)


### Comparison of run time between scripts

My orginal script ran perfectly however it took longer to run than desired. I was given an outline on how to refactor it. 
The original script ran in 0.91 and 0.92 seconds. While that is a decent amount of time, if I had a larger dataset it would take much longer to run. 

2017 runtime
![Original_2017_runtime](https://user-images.githubusercontent.com/81715217/118378170-71457a80-b597-11eb-9803-26b6b4b79c89.png)

2018 runtime
![Original_2018_runtime](https://user-images.githubusercontent.com/81715217/118378178-828e8700-b597-11eb-8b89-b79253d55d78.png)

The refactored script ran in a much quicker time!

2017 refactored runtime of 0.14 seconds
![VBA_Challenge_2017](https://user-images.githubusercontent.com/81715217/118378196-a18d1900-b597-11eb-844b-2498ea65e8ba.png)

2018 refactored runtime of 0.15 seconds
![VBA_Challenge_2018](https://user-images.githubusercontent.com/81715217/118378213-c5e8f580-b597-11eb-92e9-69ec432d6f13.png)

### The refactored part of the code

'1a) Create a ticker Index
    Dim tickerIndex As Single
        tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Double
    Dim tickerEndingPrices(12) As Double
            
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        ticker = tickers(i)
        tickerVolumes(i) = 0
                        
     Next i
     '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                                        
        
            '3d Increase the tickerIndex.
        
            tickerIndex = tickerIndex + 1
        
        End If
        'End If
              
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

## Summary

### Advantages/Disadvantages of refactoring code

Refactoring code is a good way to reevaluate your code to make it run faster. Refactoring can also make your code run even if the dataset is changed.
If you're not careful though, refactoring can create bugs that weren't there in the original code. For instance, I accidentally added an "s" to tickersIndex 
in 2 places in the refactored code that made it not run. It was time consuming debugging it. If I was on a more strict deadline that would have been very
problematic.

### Advantages/Disadvantages of refactoring this VBA code

Refactoring the code made it run quite a bit faster and therefore more effecient if per say we are give a larger dataset for 2019. 
As a stated before, a disadvantage was the 3 hours it took me to realize I had added an "s" accidentally. Another advantage is that it gave me a
deeper understanding of the thought process and language that goes with writing the code in VBA better.
