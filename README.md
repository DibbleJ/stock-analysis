# Refactoring Analysis

## Overview of Project

### Purpose

The purpose of this project is to refactor code so that the run time can be improved. This will allow for the macro to be run on larger data sets. Analysis of the refactoring will include comparison between the time required to complete analysis with the original code and the time required to complete analysis with the refactored code.

## Analysis Results

### Analysis of Refactoring

Originally, the analysis of all stocks took a little over 0.6 seconds to complete for both the 2017 and 2018 data sets. 

![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/40553064/117520565-658dfe80-af6e-11eb-8f4a-9d4367d26177.PNG)

![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/40553064/117520593-89514480-af6e-11eb-957a-9ed0e22f3a7c.PNG)

After code refactoring was completed, the analysis of all stocks took just over 0.1 seconds for both the 2017 and 2018 data sets


![VBA_Challenge_2017](https://user-images.githubusercontent.com/40553064/117520561-60c94a80-af6e-11eb-9a98-e27289f2d1e5.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/40553064/117520599-90785280-af6e-11eb-80ed-4445cbcc12f9.PNG)

The timing improvement stems from reducing the number of times that the data is analyzed by a factor of twelve! The original code below will loop through the entire data set for every ticker symbol.

'loop through tickers
For j = 0 To UBound(tickers)
        ticker = tickers(j)
    
        'set initial volume to zero
        totalVolume = 0
        
        'loop over all the rows
        For i = 2 To RowCount
            
            'increase totalVolume
            If Cells(i, 1).Value = ticker Then
                
                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(i, 8).Value

The refactored code below only loops through the data set once, and distributes the data into the appropriate array for quick distribution.

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

# Summary 
As demonstrated above, refactoring code is often necessary to optimize routines to process large data sets more efficiently. The risk in performing this refactoring is that in updating your code, errors may be introduced. There is also no guarantee the your process will run any faster than it did originally.

In relation to this script, refactoring the code has yielded completion of the analysis in almost 1/6th of the original time. While this is a large improvement, the process may make the code more challenging to interpret to future coders. This is why comprehensive commenting is crucial to good coding.

One of the primary improvements necessary to use this on a larger data set would be to generate the array of tickers from the unique values present.
