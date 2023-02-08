# Analysis of Green Energy Stocks with VBA

## Overview of the project
The main goal of this project was to help Steve and his parents with stock investments. They are heavily interested in supporting green energy. His parents are especially interested in a specific company, DAQO New Energy Corporation. However, Steve believe there should be further research to understand performance, so they can diversify their portfolio, and make an informed decision when investing. 
### Purpose 
During this analysis, the priority was to develop VBA Scirpts to analyze available data on over a dozen different stocks for the years 2017 and 2018. Showing the stock's performance meassured by trade volume, and percentage change in open/closing price while using skills within VBA creating, loops, macros, pop-ups, cell formatting, and code refactoring
## Results
Steve's parents interest in DQ is likely based on the 2017 performance, it had almost a 200% gain, however in 2018 it performed rather poorly shrinking by over 60% ENPH shows a much more promising performance over both years. 
When refactoring the intial VBA Code it runs significantly faster. 
The Initial VBA Code took about .65 seconds

Comparing it to the refactored code which took about .13 seconds
### Code
In the intial code, we pulled and output data for the volume, starting price, and ending price of one stock- resetting those variables with each iteration of the loop, 12 rounds for each ticker, and loops through the entire set of variable (3000 rows) whcich is essentially 36,000 rows of data
```
For i = 0 to 11
.....
   For j = to rowStart to rowEnd
...
   next j
next i
```
The reduced runtime was made possible by use of a tickerindex and arrays to pull and store the volumes, starting prices, and ending prices within one for loop and then outputting them to the sheet in another loop. (about 3,000 rows rather than 36,000)
```
For i = 2 to RowCount
....
next i
For i = 0 to 11
....
next i
```

## Limitations
Some disadvantages of refactoring is that it can cause errors and bugs if someone is inexperienced, or a beginner to VBA; for example, when it loops through the stocks, it assumes each is in chronological order. If the data were not in this order the results would be incorrect in a few different ways. The Volume needs the ticker symbols o be in order. For the total gains/losses the loop assumes the first instance is the earliest and the last is the latest. 
To combat these we could consider filtering the data first by ticker symbol and date, we could also add the volumes, and time stamps to take the difference factor of their close values. 
## Summary
Refactoring code is a good opportunity to improve performance and readability; make the code more extensible, and efficient; but they can take up a lot of time. The intial code was simpler since there was not arrays to worry about, however it wsa much slower.  In our particular case, refactoring code to use arrays is advantageous because it reduces our runtime. It also leaves us with array of these stored values that can be used if we choose to do further analyses or add additional code; it is cleaner to pull the information from the array versus having to retrieve the info again from within the loop. A disadvantage is that refactoring the macro may cause the script to break in some places; a change in one place will require a change in another. Another cautionary measure is saving work frequently, if moving to a significant change, if something breaks, if we have a slower code it is much more preferable to a broken code, causing further time spent on going through each line to find the cause.
