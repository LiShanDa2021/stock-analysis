# Stock Market Analysis Challenge

## Overview

A client, a graduate of finance, approached me with a stock portfolio consisting of green energy stocks to analyze. They were interested in one stock in particular -- DAQO New Energy Corporation (DQ), a photovoltaics company. The client wanted me to use my expertise with Excel and VBA to return the total daily volume for the year 2018 as well as the yearly return. Without knowledge of Excel and VBA this would have proven a mind-numbing, repetitive task and the resulting report would have included errors. However, I was able to automate the process using Excel's VBA programming language. 

### Automating Excel
In order to automate the process of summing the total yearly volume, I simply used a `For` loop to move through each row and a pair of `If` - `Then` statements to find the first and last rows of DQ's stock. My code appears in the screenshot below.

![DAQO Code](https://github.com/LiShanDa2021/stock-analysis/blob/main/daqo%20code.png?raw=true)

Of course, being a graduate of finance, my client wanted a more diverse portfolio so they included an additional 11 stocks for which they wanted the yearly volume and the yearly return. Although I could repurpose portions of the code I had written to analyze DQ, it was a more complex task and required more complex code. To solve this, I created an array to hold the 12 different stocks. I then created a `For` loop to cycle through the different stocks and a `For` loop within that loop to cycle through each of the more than 3000 rows for each stock. The code appears in the image below.

![Original All Stock Analysis Code with Nested For Loop](https://github.com/LiShanDa2021/stock-analysis/blob/main/original%20code%20nested%20for%20loop.png?raw=true)

Of course, I wanted to make the final workbook as user-friendly for my client as possible down to the salience of the results, so I created an additional subroutine that used  `If` - `Then` statements to conditionaly format the results of the analysis and then embedded that subroutine in the original macro. When a stock's yearly return was positive, the backround appeared green, and when it was negative, it appeared red (as you can see in the Results section). I even created user-facing buttons (image below) to run the analysis so my client would not have to search for the developer tab to run the macros.

![Buttons](https://github.com/LiShanDa2021/stock-analysis/blob/main/buttons%20screen%20shot.png?raw=true)

### Assessing My Code
My client was pleased with the analysis my code produced and came to me with another, considerably greater task: analyze the entire stock market. Unfortunately, my code was not up to the task of analyzing thousands of stocks. I had created a timer to see how long the program took to run. According to the timer, my program took .66 seconds to run the code for the dozen stocks for the year 2017 and .64 seconds for 2018. While hardly instantaneous, this was perfectly adequate to analyze a dozen stocks. However, it simply would not do to analyze the whole market.

### Refactoring My Code
In order to take on my client's next job, I needed better code, so I looked for weaknesses in my original. I observed that in my original script, Excel scrubbed the data from one worksheet and then wrote the data in another all in the same nested loop structure. Switching between worksheets was very likely consuming too many resources. To fix this, I removed the nested loop and instead created arrays for the data I wanted to display. Then I looped through the stock portfolio and instead of switching back to the original worksheet, I stored the data in the arrays. Then, I switched back to the report worksheet and created a short loop outside the main loop to loop through the arrays and display the data. As you will see in the results section, it worked. My code ran significantly faster.

![Refactored Code with Arrays](https://github.com/LiShanDa2021/stock-analysis/blob/main/refactored%20code%20with%20array%20loop.png?raw=true)

## Results
Here I will discuss both the results of the stock analysis as well as the increased performance of my refactored code.

### DAQO Analysis
Unfortunately for my client (or ranther, my client's parents who had high hopes for DAQO) DQ's stock price fell 63% in 2018. Given DQ's disappointing performance, it was a good thing my client handed me a diverse portfolio of green stocks.

![DAQO Performance](https://github.com/LiShanDa2021/stock-analysis/blob/main/daqo%20stock.png?raw=true)

### Green Stocks Analysis

As for Green Stocks overall, 2017 was a promising year with the entire portfolio except for TERP (TerraForm Power, Inc.) gaining value. 

![Green Stocks Performance 2017](https://github.com/LiShanDa2021/stock-analysis/blob/main/2017%20all%20stocks.png?raw=true)

2018, however, was a bust. Only ENPH (Enphase Energy Inc.) and RUN (Sun Run Inc.) gained value. The green energy market seems to be incredibly risky.

![Green Stocks Performance 2018](https://github.com/LiShanDa2021/stock-analysis/blob/main/2018%20all%20stocks.png?raw=true)

### Refactored Code Performance

Here are screenshots of how long it took to run my original code.

![2017 time to run original code](https://github.com/LiShanDa2021/stock-analysis/blob/main/time%20for%20original%20script%20to%20analyze%20all%20stocks%202017.png?raw=true)

![2018 time to run original code](https://github.com/LiShanDa2021/stock-analysis/blob/main/time%20for%20original%20script%20to%20analyze%20all%20stocks%202018.png?raw=true)

And here are screenshots of the refactored code's performace. As you can see, it is much faster at approximately .13 seconds for each year.

![2017 time to run refactored code](https://github.com/LiShanDa2021/stock-analysis/blob/main/2017%20time%20for%20refactored%20code%20to%20%20analyze%20all%20stocks.png?raw=true)

![2018 time to run refactored code](https://github.com/LiShanDa2021/stock-analysis/blob/main/2018%20time%20for%20refactored%20code%20to%20analyze%20all%20stocks.png?raw=true)

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantages of refactoring code are plain to see as show in the results. Refactored code performs more quickly thus consuming fewer resources. Additionally, refactored code is simpler and more elegant and thus easier for another programmer to understand if they need to make changes to it. Finally, there may be errors lurking in the original code that could be corrected when refactoring. 

Despite these advantages, there are reasons to forgo refactoring code. The main disadvantage is the time it requires. My code took at least an hour to refactor. Though a seasoned programmer might be able to do it faster, their time is no doubt valuable. For this reason, it is important for programmers to learn the best practices for writing code and know which design patterns are most efficient so refactoring is needed less often.

In the case of my VBA script, the choice is obvious. The original code ran at a sluggish .65 seconds. That simply unnacceptable for running an analysis on thousands of stocks. My refactored code, on the other hand, was up to the task, coming in at a fraction of the original time.
