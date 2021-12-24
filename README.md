# Writing & Refactoring Code to Analyze Stock Performance
Module 2 Challenge completed by Hannah Wikum in December 2021

## Overview of Project
### Background
This project was performed for Steve, a recent college graduate with a finance degree who is helping his parents make smarter investments in green energy. THey initially invested all of their money in DAQO New Energy Corp (DQ). Steve wants to look into the performance of DQ and figure out how to diversify his parent's funds using performance data for two years for a small number of stocks.

### Purpose
In order to analyze the data so Steve can make a recommendation on how to best diversify, I created code using nested for loops in VBA to calculate performance results, including total daily volume of trades and return over the year for both 2017 and 2018. In order to help Steve easily analyze performance for his parents, I added buttons to run the code and an input prompt to select the year so anyone can use it without VBA experience. I also added coding to change the number format and color of the output, so Steve can easily see which stocks performed well and which did not. Finally, I added a timer to check how long it took to run the code. The timer was used to compare results before and after refactoring to ensure I improved the efficiency of the code. The goal was to make the code as efficient as possible so it could eventually be used to analyze more than just 12 stocks.

## Results
### Initial Code
The first part of the analysis involved calculating the total daily volume and return for just DQ. I used VBA code to create headers for the stock name, total daily volume, and return, then found the last row of data, before getting into the for loop. The full code can be seen in my workbook, but I've highlighted the key for loop below:

![image](https://user-images.githubusercontent.com/93058069/147305593-56b52353-2e01-45b5-9e64-ad269a3cfa1b.png)

This for loop reads through every row to see if it contains data for DQ. If it does, it adds the daily volume to get a total daily volume at the end. If it is the first row that contains DQ, then it records the stock starting price. If it is the last row with ticker DQ, then it records the stock ending price, both of which were used to calculate the annual return. (Note: this is dependent on the data being grouped by ticker and data sorted oldest to newest stock price record.) The results for DQ were not good - a **-63%** return for 2018!

Since the results for DQ were not satisfactory, I modified the initial code to review the results for all 12 stocks. It used a similar structure to the code above, but added a nested for loop to review results for all tickers. I made an array of all tickers, as seen in the screenshot below.

![image](https://user-images.githubusercontent.com/93058069/147305923-ecf116bc-c9ce-4ead-8775-4bca36910eb3.png)

The outer for loop (i) now filtered through the array of all 12 stock tickers, while the inner for loop (j) calculated our three metrics described above. Here is a sampling of the for loop when updated to cycle through all tickers:

![image](https://user-images.githubusercontent.com/93058069/147306072-64128f15-c991-459b-b7f2-050ea6868ae5.png)

## Initial Findings
At the end of working on the initial code, we added an input box so any user can input the year 2017 or 2018 and a timer with message box to let us know how long it took to run the code. The code took 0.4765625 seconds to run for 2017 and 0.46875 seconds for 2018.

In 2017, all but one stock (TERP) had a positive return and DQ had the highest rate of return at 199%. In 2018, results completely changed and only two stocks (ENPH and RUN) had positive returns. DQ had the worst results at -63%.

### Refactored Code
For the second part of the analysis, I refactored the initial code to make it more efficient to run and then added formatting to improve readability. The numerical output results are identical, but the goal was to make the code run faster so it could be used for a larger data source (i.e. all stocks).

In the initial code, you have to loop through up to 3,012 rows of data once for each ticker for a total of 12 times. To make it faster, I created an index of the Tickers called tickerIndex. This allowed the code to evaluate the total volume, starting price, and ending price for all tickers at once, instead of only looking for one ticker each time it looped through the data. I also made the outputs for ticker, total volume, and return into arrays. Here is an example of using the ticker index to calculate total daily volumes instead of evaluating one ticker at a time:

![image](https://user-images.githubusercontent.com/93058069/147307194-9d03fd06-a2ff-49a4-88bc-541f87c3a810.png)

The refactored code was significantly faster than the initial code. Here is a look at the results for 2017 and 2018:

_2017_

![image](https://user-images.githubusercontent.com/93058069/147307621-8f8dac8a-8dce-4bfa-a387-e3c6af5501ee.png)

_2018_

![image](https://user-images.githubusercontent.com/93058069/147307574-40fbb240-fc4f-4ec9-83f1-8e9dee598270.png)

That is a 79% time reduction for 2017 and 80% faster for 2018, so the refactoring was successful. Also, adding the number formatting and color coding (red is a negative return, green is a positive return) greatly improved the readability of the results, as you can see below on resuls from 2017.

![image](https://user-images.githubusercontent.com/93058069/147307886-aacf50be-358a-45be-a4c7-577ad0e66e2f.png)

## Summary
Refactoring the code successfully decreased the run time per the challenge goal, but was it necessary beyond trying to get a good grade on this submission?

### Pros/Cons of Refactoring Code
There are several advantages of refactoring code. Taking the time to refactor makes your code more efficient, which is useful when dealing with large datasets. It also gives you the time to clean up the code so it is as clear and readable as possible, which would be helpful if you are intending to update the code in the future or want to share with others. On the negative side, it can take a long time to refactor (if it is even possible), which would be a waste of resources if your initial code produced the results you needed and you don't intend to reuse the code.

### Pros/Cons of Initial versus Refactored Code
For our dataset of annual data for 12 stocks, the initial code was sufficient to analyze the data and provide results quickly. The pro was that it was faster for me to build than the refactored code and I have an easier time understanding it even when I go back several days later, but the con is it would take a lot longer if you tried to use it to analyze results for more stocks, especially if you tried to look at the performance of the entire stock market. On the other hand, the advantage of the refactored code is that it does run faster, but the disadvantage is it is much more complex to build and understand. For this purpose, it took me significantly longer write the refactored code than time saved when it ran; I would have to run the code approximately 29,000 times before I would have recognized the time savings.

