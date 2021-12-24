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

![image](https://user-images.githubusercontent.com/93058069/147300720-b8659138-3793-47a4-95d5-496cd96aed68.png)

This for loop reads through every row to see if it contains data for DQ. If it does, it adds the daily volume to get a total daily volume at the end. If it is the first row that contains DQ, then it records the stock starting price. If it is the last row with DQ, then it records the stock ending price, both of which were used to calculate the year return. (Note: this is dependent on the data being grouped by ticker and data sorted oldest to newest record.) The results for DQ were not good - a **-63%** return for 2018!

Since the results for DQ were not satisfactory, I modified the initial code to review the results for all 12 stocks.

### Refactored Code
