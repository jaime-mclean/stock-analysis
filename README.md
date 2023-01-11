# Green Energy Stock Analysis

---
## Project Overview

---
Stock Market analysis can be simplified if you are looking at annual trends for a given stock or market segment. The provided worksheet will analyze the stocks provided for the years included, and can be expanded to other years if desired. Note that sheets for additional years will need to formatted identically to the existing worksheets, with the data sorted by ticker then by date, and the worksheet will need to be named with the 4-digit year. The code has been optimized, so if additional stocks are added to the analysis it should perform well. The ticker symbols will need to added and the array expanded, but that is a simple adjustment to the code.

---
## Results

---
### Analysis of twelve stocks

Green Energy is a rapidly expanding market segment. Not unexpectedly, the valuation of companies in emerging markets can be  volatile. Using the analysis worksheet I have provided, the data show that most companies showed aggressive growth in 2017. All but two of the stocks listed were plagued with losses, significant in some cases. Per this analysis, **ENPH** and **RUN** appear to have sustained growth in 2018 on top of their 2017 growth (see graph below).   [Return by Year](https://github.com/jaime-mclean/stock-analysis/main/Resources/ReturnByYear.png)

Volume for both **ENPH** and **RUN** remained high over both 2017 and 2018, with 2018 seeing higher volumes traded than 2017. The correlation between the volume traded and teh stock price is weak, as some of the poor performing stocks (notably **DQ**) had a higher trading volumes in 2018 (see graph below). [Volume Traded by Year](https://github.com/jaime-mclean/stock-analysis/main/Resources/VolumeTradedByYear.png)

---
### Analysis Worksheet

The provided worksheet contains an optimized VBA macro that can be used on the existing data or additional data as described in the Overview. The code optimizations include consolidation of the volume and Return statements from multiple loops to a single loop, followed by outputting the data into the table and formatting the return for each stock (green for growth, red for contraction) for ease of interpretation. The optimized code, shown below, iteratively adds the daily volume for each stock (based on ticker symmbol) and determines the starting price and ending price based on the first and last entry for each stock. This is why the data must be sorted by ticker and then by date, as failure to do so may provide an incorrect starting or ending price and a subsequent incorrect return. 

` For tickerIndex = 0 To 11
        
        tickerVolumes(tickerIndex) = 0
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
            If Cells(i, 1) = tickers(tickerIndex) Then
        
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            End If
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
            '3c) check if the current row is the last row with the selected ticker
        
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
            'If the next row's ticker doesn't match, increase the tickerIndex.
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            End If
    
        Next i
        
    Next tickerIndex
`
The original code ran slightly slower than the optimized code, as can be seen by the run times provided each time the code is run. The initial code run times shown in the images below indicate that the run times were already quite fast. 
[2017 initial Run time]
[2018 initial Run time]

With the addition of more stocks, the run times could become increasingly long. THe optimized code runs faster by a factor of 10 (or more), and should easily  support additional input. The timers for the optimized code are shown below.
[2018 refactored run time]
[2018 refactored run time]

---
## Summary

---
As discussed, the refactored code is much faster than the original code. The original code was slowed by the multiople loops, and with an increase in speed is a loss of modular portability. Since all three loops have been combined into a single loop, there is not an easy way to port a single portion of the code into a new worksheet. This can be easily done with the original code, but that portability also takes time that is reflected in the run time. Although the difference in run time seems inconsequential with the handful of stocks analyzed here, the optimation will pay off with larger data sets.

---
