# Comparing the Volume and Yearly Rates of Return for a List of Green Stocks Using VBA Analysis 
Columbia Bootcamp VBA project
## Results
###### Stock performance
Compare performance between 2017 and 2018. Suggest broadening out to a larger sample of stocks to find more competitive securities for the client's portfolio. In order to handle more stocks and more years of data without increasing the run time of the analysis too much we needed to refactor the code.
###### Performance of the code
We looped through the data once rather than the 4x in the original code. Then we placed the outputs on the output sheet one time instead of at the end of each loop. (Add images showing old code and new code of loop) This saved X amt of time in the run time. (Add images of old runtimes and new runtimes)
## Summary of refactoring
###### The pros and cons of refactoring code
The disadvantage of refactoring the code is that it increases coding time, introduces the chance of adding errors, but the result is a more flexible script that can be adjusted to add tickers and additional years of data with only very small changes in the array and by adding and activating new sheets.
###### Challenges of refactoring this VBA project's script
The advantages of refactoring the code is that it can now run faster and handle a larger data set, including new stocks and more years of data. The original code also did the job but the new code can be run faster and is more efficient so that it can be more easily expanded to analyze a larger data set.

