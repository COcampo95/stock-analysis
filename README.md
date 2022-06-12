# stock-analysis

## Overview of Project

### Purpose
This project was to refract the Microsoft VBA coding for Steve's parents' annual stock performance from 2017 and 2018 to determine which environmental energy resource stocks are worth investing in. 

### Results 

#### VBACode
Before altering the code, I copied the existing code from the previous macros when calculating the yearValueAnalysis and altered the code to the specifications for this project. Below is the refracted code:

 Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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

        '3b) Check if the current row is the first row with the selected tickerIndex.      'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
            '3d Increase the tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
            
        End If
            
            
        'End If
    
Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
       
    Next i
    

#### Stock Performances 
Comparing 2017 and 2018, EMP and RUN continue to perform well in the market. Although they do not yield as high of a return, they continue to do well in a market that saw all other stocks drop in performance. 

![All Stocks Analysis_2017](https://user-images.githubusercontent.com/105950742/173254351-e9e21b35-374e-4b72-8a9c-ab4cd1372004.png)

![All Stocks Analysis_2018](https://user-images.githubusercontent.com/105950742/173254357-b9b93c48-22f3-4c1e-a3be-14445c0f79ac.png)


### Summary

#### Pros and Cons of Refactoring code

The advantages of refactoring code are the is the code is more organized and cleaner. Refactoring the code decreased the macro run time; prior to the changes, the code took about 1 second to run. 

![All Stocks Analysis_2017_Time](https://user-images.githubusercontent.com/105950742/173254496-5e6fbcf8-551b-435e-9409-acefd1e72e18.png)

![All Stocks Analysis_2018_Time](https://user-images.githubusercontent.com/105950742/173254502-dd30a625-6f95-47cf-8e64-80c0180ec06c.png)

The disadvantage to refactoring is if the application is too large to take the proper time to alter the code and run proper testing. 
