# Writing & Refactoring Code to Analyze Stock Performance
Module 2 Challenge completed by Hannah Wikum

## Overview of Project
### Background
This project was performed for Steve, a recent college graduate with a finance degree who is helping his parents make smarter investments in green energy. THey initially invested all of their money in DAQO New Energy Corp (DQ). Steve wants to look into the performance of DQ and figure out how to diversify his parent's funds using performance data for two years for a small number of stocks.

### Purpose
In order to analyze the data so Steve can make a recommendation on how to best diversify, I created code using nested for loops in VBA to calculate performance results, including total daily volume of trades and return over the year for both 2017 and 2018. In order to help Steve easily analyze performance for his parents, I added buttons to run the code and message box prompts to select the year so anyone can use it without VBA experience. I also added coding to change the number format and color of the output, so Steve can easily see which stocks performed well and which did not. Finally, I added a timer to check how long it took to run the code. The timer was used to compare results before and after refactoring to ensure I improved the efficiency of the code. The goal was to make the code as efficient as possible so it could eventually be used to analyze more than just 12 stocks.

## Results
### Initial Code
The first part of the analysis involved calculating the total daily volume and return for just DQ. I used VBA code to create headers for the stock name, total daily volume, and return, then found the last row of data, before getting into the for loop. The full code can be seen in my workbook, but I've highlighted the key for loop below:


### Refactored Code
