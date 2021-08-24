# An Analysis of Stocks Using VBA
## Overview of Project

### Purpose 
The purpose of this project was to edit, or refactor, the code created in Module 2 in order to collect the dataset to include the entire stock market over the last couple of years (2017 and 2018). This will help determine whether refactoring the code will make the VBA script run faster or not. The code worked very well for a few stocks, but we are interested to see if it will work well for thousands of stocks in an efficient time.
### Background
The goal of this project is to analyze a numerous amount of green energy stocks to analyze which stocks are worth investing in versus which ones are not. Using VBA to see the data of the stocks decreases the chance of errors in the results. I am analyzing twelve different stocks which will tell me the "Total Daily Volume", and the "Return" on those stocks, which are color coordinated to visually see which stocks are worth investing in. 
## Results

### Analysis of Stock Performance

Here are the results for the 12 stocks in 2017

![Analysis_2017](https://user-images.githubusercontent.com/88408350/130680551-74f2f892-6f0b-485a-a7bb-c31a1ccfb39b.PNG)

Here are the results for the 12 stocks in 2018

![Analysis_2018](https://user-images.githubusercontent.com/88408350/130680602-dbe12023-d168-45ba-a1ff-28afd2ae4e99.PNG)

Using this code below, I was able to get the output for "Total Daily Volume" and "Return" for each stock. We are able to see that the stock performance was better in 2017 than 2018. Only one stock in 2017 (TERP) had a negative return while 10 stocks in 2018 had negative returns - which are highlighted in red. 

    '1a) Create a ticker Index
        tickerIndex = 0'1b) Create three output arrays
        
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
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        
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

### Analysis of Execution Times

Here is the execution time of the refactored script in 2017 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/88408350/130681825-e8053bc9-d55d-4f2c-a138-4278826d59b3.PNG)

Here is the execution time of the refactored script in 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/88408350/130681840-ca8304dd-12f5-44d5-b73c-04aac9b8e089.PNG)

In our original script, it took aboput 1 second to run our analysis. After refactoring our code, it took about 10 times faster to run. This proves that refactoring our script does improve efficiency time.

## Summary

- **What are the advantages or disadvantages of refactoring code?** The advantages of refactoring code is decreasing the amount of time it takes to run our analysis. With a really massive dataset, it could take a long time to get our results, but this helps improve efficiency. It also allows us to visually see our analysis in a way that is easy to read for anyone looking at the data. A disadvanatge of refactoring code is it can be time consuming. If there is a deadline approaching, it might not be worth it to refactor, especially if the code is running smoothly.

- **How do these pros and cons apply to refactoring the original VBA script?** These pros and cons of refactoring apply to our original VBA script because we were able to get the results we were looking for in a clean manner that took one-tenth of a second to get. However, this took a lot of planning and research to achieve, even though we got the same results before refactoring our code. 
  
