# Using VBA to Compare the Performance of a List of Stocks 
## Purpose of the project
###### To analyze the performance of 12 green industry stocks in 2017 and 2018 to determine the volume they traded at and their rate of return in both years, using VBA macros.
## Results
###### Stock performance
Our analysis shows that only two stocks in the client's set have positive returns for 2017 and 2018, ENPH and RUN. Based on this performance I would suggest the client only invests in ENPH and RUN out of this set of potential stocks. ![Green Stocks 2017-2018 Performance](Images/2017-2018Performance.png)
 To find more promising candidates for the client's portfolio, the same analysis could be performed on a larger sample of stocks to identify more competitive stock names.

###### Performance of the code
 In order to be able to run the same analysis on a greater number of stocks for the client in future without increasing the run time, it was necessary to refactor the code to make it run more efficiently. In order achieve this we re-wrote the script to use a tickerIndex to loop through the data once, rather than the multiple times needed in the original code. ![Original Code Loops](Images/CodeComparison.png)

## Summary of refactoring
###### The pros and cons of refactoring code
###### Advantages
In this case, refactoring the code saved .27 seconds or 77% in run time on the 2017 sheet with similar results on the 2018 data. ![Refactored Code Loop](Images/Times2017.png) 
###### Potential Disadvantages
The downside of refactoring the code is that doing so adds coding time and increases the chance of introducing errors, but the result is a more flexible script that can be adjusted to add tickers and additional years of data with only small changes in the code required.

###### Challenges of refactoring this VBA project's script
In this project,  refactoring the code made it more flexible so that it can now run faster and handle larger data sets, including new stocks and more years of data without adding run time. As the code is now any new sheets containing new tickers would necessitate manually updating the array of ticker names in the code. But in future, the code could be refactored to make it possible for the macro to automatically adjust the array of tickers so that new sheets of data could be analyzed without needing to manually increase the number of tickers and re-name the array of tickers in the code. This would make the macro we've created a much more flexible tool for analyzing a larger set of stocks with less coding hours needed.

