# Module 2: VBA of Wall Street Challenge

## Overview of the Project:

Our client, Steve recently graduated with his finance degree and his parents would like to be his first clients. They are very passionate about green energy and have their hearts set on investing in DAQO New Energy Corp (ticker: DQ). After analyzing how DAQO’s stocks fared in 2017 and 2018, Steve has recommended that his parents consider investing in other companies. To help diversify his parent’s portfolio, Steve asked for our support creating a VBA script that can quickly search through the stock results of all the green energy companies, compile the total volumes and return percentages for each company, and format the results to quickly draw attention to the companies with positive and negative return percentages.

Steve was so pleased with the VBA script we helped him build to analyze green energy companies, he would like us to refactor our code so that he can use this subroutine on the entire stock market’s results. While our previous script worked well on about a dozen companies, using it to analyze thousands of companies’ stock results, over various years will be quite cumbersome. Our task is to decrease our runtime and simplify our code so that when this subroutine is applied to the entire stock market it runs as efficiently as possible.

## Results

A good portion of the script we have already written can be borrowed with little to no modification necessary to get us started. The AllStocksAnalysis() subroutine already sets up the timer functionality, the year input box, headers, initializes the tickers array and gets the last number of rows within a worksheet. 

### This portion of the script is already provided and does not require refactoring.

### 1a.) For the purposes of refactoring, we begin by creating a tickerIndex which is the variable we will use to access all the indexes across multiple arrays. Since the tickerIndex will be used to iterate through rows of data, we need to declare it and set it to 0 outside of any conditionals.

    Dim tickerIndex As Integer
    tickerIndex = 0

### 1b.) Next we have to declare 3 new, empty arrays; tickerVolumes, tickerStartingPrices, and tickerEndingPrices. We set the data type for tickerVolumes to Long and the data type for the ticker starting and ending price arrays to Single. Since Steve intends to use this subroutine in the future on data from the entire stock market, it makes sense to ensure these are dynamic arrays given the amount of stock records will always vary. We have already borrowed code from the internet that allows us to find the last row of an entire worksheet and have named the variable RowCount. Our arrays have already been declared, so we can use the ReDim statement to change the dimension of the array based on the variable RowCount.

    Dim tickerVolumes() As Long
    ReDim tickerVolumes(RowCount)
    Dim tickerStartingPrices() As Single
    ReDim tickerStartingPrices(RowCount)
    Dim tickerEndingPrices() As Single
    ReDim tickerEndingPrices(RowCount)

### 2a.) After declaring the new arrays we have to set tickerVolumes to 0 for each of the tickers. Previously, we used a nested for loop followed by a few if statements. Nesting multiple conditionals slows down the run time; before each tickerVolume can be set to 0, it must loop through the entire code block within the larger conditional. We could create a standalone for loop to set tickerVolumes to 0; however, there is an even more efficient way to achieve this. Since each of the ticker’s tickerVolumes is being set to 0 and because we have already ReDim’d the variable, we can simplify this to one line of code: 

    tickerVolumes(j) = 0

### 2b.) Now we begin our nested conditional which we can make both dynamic and more efficient. We can borrow the code from our previous subroutine to begin the for loop. By using the RowCount variable in the for statement, we ensure that regardless of how many data records are analyzed, the script will know to execute until the last row.

### 3a-3d.) Next we need to step through the nested if statements. These three if statements are used to increase the volume of the current ticker, check if the current row is the first row that corresponds to the selected tickerIndex, check if it is the last row, and then if the following row’s ticker doesn’t match, increase the tickerIndex.

### 4.) The final step in the bit of code we are refactoring is to tell the script where, and how to output the results.

### Lastly, we format the output, end the timer and end the subroutine.

## Stock Analysis Results

Overall, we can tell that 2017 was a much better year for the green energy companies then 2018 was. DAQO was certainly not alone, in terms of disappointing returns but it did have the most drastic swing from up nearly 200% in 2017 to down nearly 63%. Only having been provided two years’ worth of data, its nearly impossible to draw any solid conclusions but it would be worth determining how volatile investing in the green energy sector really is. If Steve’s parents have their heart’s set on investing in green energy, they may want to consider investing in ENPH and possibly RUN as well. Both of these companies saw a positive return in both 2017 and 2018. In fact, even when nearly every other green energy sector saw a return loss in 2018, ENPH still saw nearly 82%.

## Run Time Results

After refactoring the script, I was able to cut down on the runtime significantly. Below are graphics displaying the in the run time results both pre and post refactoring, for each year.

## Summary

In summary, the process of refactoring code allows us to process data more efficiently without changing the desired outcome. The more we streamline code, optimize opportunities to make our code dynamic, and cut down on redundancies and nested processes, the more scalable our code becomes. In addition to the above-mentioned benefits, the process of refactoring code can also help us to learn even more techniques and skills by breaking apart complicated pieces of script and attempting to simplify the process as much as possible. One of the disadvantages of refactoring could be that changing code, simply for the sake of code, does not always result in more efficiency. One way we can verify what we are changing is improving functionality is to use the run time process. In our VBA script, we can see that the process of refactoring does in fact make the code more efficient. The main disadvantage is that refactoring code can take a long time do, for only a couple seconds worth of saved time (assuming the subroutine runs without any additional errors).

## Sources Consulted

https://www.techopedia.com/definition/3865/refactoring

https://rubberduckvba.wordpress.com/tag/refactoring/

https://www.automateexcel.com/vba/declare-dim-create-initialize-array/

