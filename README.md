# AllStocks_analysis


## Overview of Project
The project is to write readable code to find Total Daily Volume in 2017 and 2018 and sum up to give an idea of the yearly total volume and how often  it gets traded to the clients.  Then using VBA script to create a run button to run the code of analysis so they can use one button and put in the year to get the idea.
### Purpose
Using VBA script to set a nested loop and for a loop to run the data. We also use different code to find a way to let excel run the code faster to show large data set's analyzing results. 

## Analysis of All stocks between 2017 and 2018
### Analysis for VBA_Challenge_2017
![](http://https://github.com/Daisyzhao21/AllStocks_analysis.git)
The data set of All stocks 2017 (VBA_Challenge_2017.png) show that the eleven company's name in the column A; total daily trading amount in 2017 in the column B; and yearly return of each company in column C. We can see except for the ticker TERP has a negative retun -7.2%, all the other tickers have a positive return range from 5.5%to 199.4%.

We also use for loop in refactored scrip to increase the running speed, making separate loops instead of a nest loop which run all the variable by steps for a large number of times. The original running speed is 5.976563 second and refactored script's running takes only 1.367188 seconds. 
### Analysis for VBA_Challenge_2018
![](https://github.com/Daisyzhao21/AllStocks_analysis.git) 
The data set of All stocks 2018 (VBA_Challenge_2018.png) show that the eleven company's name in the column A; total daily trading amount in 2018 in the column B; and yearly return of each company in column C.  We can see except for the ticker ENPH and RUN have return of 81.9% and 84.0%, all the other tickers have a negative return range from -3.5% to 62.6%,so, year of 2018 most company have a negative returns. 

Same as the nested loop in the original script, it takes 3.914063 seconds for data 2018 and the refactored script takes only 1.367188 seconds.


## Summary
### The advantages and disadvantages of refactoring code 
The most obviously advantage is lessen the time to run the script for the computer by using single loop instead of a nested loop which a loop inside another loop. The disadvantage of the code is that the code become more complex and consists of significantly more code.

### The advantages and disadvantages of the original and refactored VBA script
 The original code is easy for people to write because we can let the computer to deal with conditions itself. We just give a loop and put another loop under the first loop. However, the refactored script requires more complex think of dealing with different variables, which will increase the thinking process of putting all the data together in a correct order.

On the other hand,the original script need longer time to process data from a loop inside a loop which is a disadvantage especially for large data sets. The refactored script increases the process speed to make sure the computer have smaller scale of calculating process.
