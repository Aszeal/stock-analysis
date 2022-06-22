# Stock Analysis with Excel Visual Basic for Applications 

## Overview of Project

### Purpose of Project

Tasked with the creation of a Macro in *Visual Basic for Applications* for a client named Steve. Steve wanted a Macro he could use to analyze data of stocks for multiple years. Steve wanted a Macro that was fast, reliable, and easy to use for the years 2017 and 2018. After all, he wanted this Macro to help his family make better invesment decisions. But in order to accomplish this goal I needed to refactor the code and determine which script preformed the most efficiently. 

## Data and Modeling Approach 
The data that I am presenting includes tables, screentshots, and a visual of the refactored code.


The two tables(Charts) contains the following information:
1. Stock information that includes 12 different stocks 
    ( Ticker value 
    , Date the Stock was issued 
    , Starting and Ending Prices 
    , Volume)
    
2. Toal Daily Volumne 

3. Return 

The refactored code is showcased to show what imporvements were made to increase efficiency between the original script and the refactored script. Lastly, the screenshot shows the differences between the execution times of the year 2017 and 2018.

## Results
### Analysis
Refactored Script:
In order to refactor the code I was provided with an alternative code that contained useful script information to make adjustments to run the code faster. Also, the code provided the necessary information to create an input box that shows the execution times of the original script and refactored script. Below, is the the completed refactored script. 




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
    


Tables(Charts):
 
 This is the table that displays the information for the stock year 2017.
 
 https://github.com/Aszeal/stock-analysis/blob/main/Resources_VBA_Challenge/All%20Stocks%202017.png
 
 This is the table that displays the information for the stock year 2018.

https://github.com/Aszeal/stock-analysis/blob/main/Resources_VBA_Challenge/All%20Stocks%202018.png



## Summary

### Avantages and Disadvantages of Refactoring Script
Refactoring is something that can be very beneifical for the efficiency of VBA. Usually, refactoring should be utilized when your code is completely finished and runs correctly. Only then, can you conduct refactoring. Refactoring serves many purposes such as, improving the speed of execution of code, more clearer and concise code, and ease-of-use. Now, refactoring can be harmful if your code is not correctly scripted. Users will try to refactor a clients code, but will utterly fail because the original code is buggy; resulting in the creation of more bugs in the code. In conclusion, refactoring is a powerful tool that should be utilized once your script is running properly.


### Advantages and Disadvantages of The Original and Refactored VBA Script
The advantages for the refactorization of the VBA script is quite overwhelming. Most importantly, it decreased  the run-time more than 50 percent. Based off my research this is significant. When dealing with much larger sets of data execution time is everyting. It determines how quickly you can submit a report for a business or complete a project for a client. In this case, Steve will be happy to know that the refactored code not only runs faster he can input different stock years, and get the information he needs. One disadvantage I ran into is how precise everything needs to be. Without proper notation, remembering what the code does can become very bad very quickly. In conclusion, refactoring script is a great tool if you remember how your code works.

Below is the two screenshots of the execution times of the year 2017 and 2018:


