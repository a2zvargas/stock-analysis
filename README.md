# Stock Analysis

## Overview of Project

This project performs data analysis on stock market data to get total daily volume and yearly return percentage of a set of stocks.

## Results

### Stock Performance
Looking at the results of the stock analysis we see that almost every stock performed better in 2017 than in 2018.  All stocks but TERP had a positive rate of return over the year in 2017, but only stocks ENPH and RUN had a positive rate of return in 2018.  Although ENPH had a positive rate of return in 2018 at 81.9%, it had an even higher rate of return in 2017 with rate of return of 129.5%.  However, RUN improved from a 5.5% rate of return in 2017 to a 84% rate of return in 2018.  While TERP improved from a -7.2% rate of return in 2017 to a -5.0% rate of return in 2018 it remains a negative return.

2017 Stock Performance | 2018 Stock Performance
:-------------------------:|:-------------------------:
![stock performance 2017](https://user-images.githubusercontent.com/85706721/125175655-93283b80-e19b-11eb-8bdd-c564c17af412.png)|![stock performance 2018](https://user-images.githubusercontent.com/85706721/125175657-97545900-e19b-11eb-8278-098203c02068.png)

### Script Performance

By refactoring the script for the 2017 data we saw a 3% reduction in the length of time the script took to run, and a 6% reduction for the 2018 data.

Original Script Run Time for 2017 data | Refactored Script Run Time for 2017 data
:-------------------------:|:-------------------------:
![Original_2017](https://user-images.githubusercontent.com/85706721/125175839-f5ce0700-e19c-11eb-95ab-3e342fd7a799.png)  | ![VBA_Challenge_2017](https://user-images.githubusercontent.com/85706721/125175843-fa92bb00-e19c-11eb-8ae9-adf901aee578.png)

Original Script Run Time for 2018 data | Refactored Script Run Time for 2018 data
:-------------------------:|:-------------------------:
![Original_2018](https://user-images.githubusercontent.com/85706721/125175845-fd8dab80-e19c-11eb-93c0-aa25eb556117.png) | ![VBA_Challenge_2018](https://user-images.githubusercontent.com/85706721/125175847-ff576f00-e19c-11eb-9351-7abc8e94a2c3.png)

Storing the results in arrays during the loop instead of outputting it to the worksheet every time contributed to the performance gain.
```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

Once we had all the data in the arrays, we used a loop to output the data to the worksheet.
```
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Next i
```
## Summary

### What are the advantages or disadvantages of refactoring code?

The advantage of refactoring code is that you can improve the performance of it.  While in this exercise we were only working with a few thousand rows of data and only saw an almost negligible reduction in run time, had we been working with a larger data set it could have been a bigger reduction.  In more extreme cases the code may break with large data sets and refactoring would be required.

The disadvantage of refactoring code is that it can be time consuming, and in some cases done is better than perfect.  On the job with other project deadlines looming, one may not have the time to refactor the code.

### How do these pros and cons apply to refactoring the original VBA script?

The advantage of refactoring code to improve performance applies to this project because we improved performance.  Although it was a minimal gain, it was a good learning experience.

The disadvantage of refactoring code being time consuming was also present here, but again it was a good learning experience and worth the time it took. 


