# Comparing the Performance of Stocks Using VBA
## Purpose of the project
###### To analyze the performance of twelve green industry stocks in order to compare the volume they traded at and their rate of return in 2017 and 2018 with the aim of forecast which stocks are most likely to increase in value.
## Results
#### Stock performance
My analysis shows that only two stocks in the client's set have positive returns for 2017 and 2018, ENPH and RUN. Based on this observation I would suggest the client only invests in ENPH and RUN out of this set of potential stocks. ![Green Stocks 2017-2018 Performance](Images/2017-2018Performance.png)
 To find more promising candidates for the client's portfolio, I would recommend broadening out the search by applying the same analysis on a larger sample of stocks to identify additional competitive candidates.

#### Performance of the code
 In order to be able to run the same analysis on a greater number of stocks in future without increasing the run time, it was necessary to refactor the code to make it run more efficiently. In order to achieve this, we re-wrote the script to use a tickerIndex to loop through the data once, rather than running the loop multiple times to capture each data point for each ticker symbol as was done in the original code. ![Comparing Original and Refactored Code Loops](Images/CodeComparison.png) *Code loops presented here, side by side for comparison, with the older code on the left and the refactored code on the left.

## Summary of refactoring
#### The pros and cons of refactoring code
###### Advantages
In this case, refactoring the code saved .27 seconds or 77% in run time on the 2017 sheet with similar results on the 2018 data. ![Refactored Code Loop](Images/Times2017.png) 
###### Potential Disadvantages
The downside of refactoring the code is that doing so adds coding time and increases the chance of introducing errors, but when successful the result is a more flexible script that can be adjusted to analyze new sets of data with only small changes in the code required.

###### Challenges of refactoring this VBA project's script
In this project, refactoring the code made it more flexible, enabling it to run faster and handle larger data sets, including new stocks and more years of data without adding run time. As the code stands now, analysing new sheets with new tickers would necessitate manually updating the array of ticker names in the code. But in future, the code could be refactored to make it possible for the macro to automatically adjust the array of tickers so that new sheets of data could be analyzed without needing to manually increase the number of tickers and re-name the array of tickers in the code. This would make this macro a much more flexible tool for analyzing larger sets of stocks automatically by the client with less coding hours needed.

