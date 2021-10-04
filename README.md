# Stock Analysis With VBA
## Overview of Project
 The purpose of this project was to refactor code with VBA to more efficiently run an analysis on stock data from 2017 and 2018 to determine which stocks are worth investing in. The dataset included information from 12 different stocks including the ticker ID, the date the stock was sold, the opening and closing prices, the highest and lowest prices and the volume of the stock for both 2017 and 2018. From the raw data the total daily volume and the return was obtained for each stock. The code had already been created to obtain these findings during a previous module, the goal of this project was to edit the VBA script be able to run it faster for both years. 

## Results
### Stock Performance Comparison Based on Year
From on the table formulated from the VBA script that was written, it can be concluded that the stocks performed better in 2017 than they did in 2018. With the exception of the stock TERP, all other stocks had a positive return in 2017. However, in 2018 all stocks besides ENPH and RUN had negative returns.  

### Code Execution Time 
During the refactoring process, the code was edited and condensed in order for it to run quicker and more efficiently. The part of the code that was most refactored pertained to obtaining the total daily volume and return output from the 2017 and 2018 raw data. 

![Refractored Code](https://github.com/mayamtims/Stock-Analysis/blob/main/Resources/refractored_code.png)

Code to record the time in which it took the analysis to run was also included to see which version of the code ran faster. As seen in the images below, the refactored code ran quicker than the original code for both years. The images on the left show the time in which it took for the original code to run and the images on the right show the time in which it took for the refactored code to run. 
![2017 original code](https://github.com/mayamtims/Stock-Analysis/blob/main/Resources/Original_VBA_Challenge_2017.png)
![2017 refactored code](https://github.com/mayamtims/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018 original code](https://github.com/mayamtims/Stock-Analysis/blob/main/Resources/Originalz_VBA_Challenge_2018.png)
![2018 refactored code](https://github.com/mayamtims/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
### Advantages and Disadvantages of Refactoring Code
Refactoring code reduces the amount of memory and steps required to write and run the code. It increases efficiency by allowing the code to run quicker and simplifies the steps for future readers and users of the code. Disadvantages of refactoring code include the fact that the process could be time consuming, and running into code errors or confusion where there wasn't any before. 

Pertaining to this specific project, refactoring the code with VBA was adventageous because it reduced the amount of steps and time it took to run the code, making it much more efficient. A disadvantage to refactoring this code was that knowing where and how to condense and combine code lines was at times confusing and it took awhile to piece it all together in a more functional way. 