# Stock Analysis with Excel Visual Basic for Applications 

## Overview of Project

### Purpose of Project

Tasked with the creation of a Macro in *Visual Basic for Applications* for a client named Steve. Steve wanted a Macro he could use to analyze data of stocks for multiple years. Steve wanted a Macro that was fast, reliable, and easy to use for the years 2017 and 2018. After all, he wanted this Macro to help his family make better invesment decisions. But in order to accomplish this goal I needed to refactor the code and determine which script preformed the most efficently. 

### Data and Modeling Approach 
The data that I am presenting includes tables, screentshots, and a visual of the refactored code.


The two tables(Charts) contains the following information:
1. Stock information that includes 12 different stocks 
    ( Ticker value 
    , Date the Stock was issued 
    , Starting and Ending Prices 
    , Volume)
    
2. Toal Daily Volumne 

3. Return 

The refactored code is showcased to show what imporvements were made to increase efficiency between the original script and the refactored script. Lastly, the screenshot shows the differences between the run-times of the year 2017 and 2018.



### Results 

#### Analysis 
In order to refactor the code I provided with an alternative code that contained useful script information to make adjustments to run the code faster. Also, the code provided the necessary information to create an input box that shows the execution times of the original script and refactored script. Below, is the the completed refactored script. 




    tickerIndex = 0

    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    
    For i = 2 To RowCount
    
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date

### Analysis of Outcomes Based on Goals

### Challenges and Difficulties Encountered

## Results

