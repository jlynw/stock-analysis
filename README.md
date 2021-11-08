# stock-analysis
## Project Overview
The purpose of this project is to assist Steve with analyzing a dataset consisting of different green energy stocks. This is be accomplished by using VBA to automate the analyses. 


####  Initial Code
The initial code below was used to find the total daily volume and the return of twelve tickers in the years 2017 and 2018.

![initialcode1.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/initialcode1.PNG)
![initialcode2.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/initialcode2.PNG)


Indicated below is how long it took the code to analyze the data and compile the results.
![initialVBA_Challenge_2017.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/initialVBA_Challenge_2017.PNG) ![initial_VBA_Challenge_2018.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/initial_VBA_Challenge_2018.PNG)


Below are the total daily volume and the return of tickers in 2017 and 2018. The values highlighted in green shows a positive return, while values highlighted in red shows a negative return. The majority of stocks performed better in 2017 compared to 2018. ENPH and RUN were the only tickers that did not have a negative return.

![Stock_Analysis_Output_2017.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/Stock_Analysis_Output_2017.PNG)
![Stock_Analysis_Output_2018.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/Stock_Analysis_Output_2018.PNG)

#### Refacored code
Below is the refactored portion of the initial code above. This enabled the code to execute faster. First, a tickerIndex variable was created. Then, three output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) were assigned. Next, a script was written to add the ticker volume for each ticker. Two if-then statements used to check the current row for each index and assign the starting and closing prices to the tickerStartingPrice and tickerEndingPrice variables. Lastly, a for loop was written to the four arrays (tickers, tickerVolumes, tickerStartingPrice, tickerEndingPrice) to give us the outputs in the All Stocks Analysis worksheet. In conclusion, using the variables: tickerIndex, tickerVolumes, tickerStartingPrice, and tickerEndingPrice enabled the code to execute faster than the original code. 

![refactored.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/refactored.PNG)

Indicated below is how long it took the code to analyze the data and compile the results.
![VBA_Challenge_2017.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG) ![VBA_Challenge_2018.PNG](https://github.com/jlynw/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary
The advantages of refactoring is that it can make your code more efficient and run faster. A disadvantage is that refacoring can introduce runtime errors or bugs to a perfectly functioning code.

In the case of the stock analysis, an advantage of refacoring significantly decreased from an average of 0.70 seconds to 0.13 seconds. A disadvange is that I ran into run time errors 9 (out of range subscript) and 6 (overflow). This occured when I was assigning the new variables tickerIndex and when I was assigning the datatype for the tickerVolumes variable.
