# VBA_Challenge
## Overview of Project
This project was intended to assist Steve in finding the total daily volume and total daily return for multiple stocks so that Steve's parents can be informed on which stocks they should invest in. The total daily volume is the total number of shares that are traded throughout the day so that we can measure how quickly a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. We used VBA, a tool within Micrsoft Excel, to write code that will find the total daily volume and total daily return. The second part of this project was to refactor the VBA code that was originally used in the tutorial so that it runs quicker and more efficiently. Though the original code works for a dozen stocks, if Steve would like to use this code in the future for a larger amount of stocks we will need the code to run quicker, take fewer steps and essentially accomplish the task more simply.  

## Results

### Stock Performance
![2017_All_Stocks_Return](https://user-images.githubusercontent.com/85354946/156414477-eda34417-2457-4960-9826-5e9602e47618.png)
![2018_All_Stocks_Return](https://user-images.githubusercontent.com/85354946/156414494-debb7133-eb90-438a-acdc-453fe6414451.png)

Overall, almost all the stocks in 2017 had positive returns with the exception of TERP. The highest performing stock for 2017 was DQ with about a 199% return making it look like a desirable option for Steve's parents to choose in that year. However, DQ however had the biggest change over the years, starting out strong in 2017 and then moving into a 60% loss in 2018. Looking at the 2018 stocks, a majority of these same stocks are performing much worse with negative returns. Though ENPH took a large return cut in 2018 as compared to 2017, it still kept a positive return. RUN, who had a 5.5% return in 2017 now had a 84% return in 2018 showing its growth throughout the year. Some of the lowest performing throughout both 2017 and 2018 were TERP and AY as TERP was at a lost both years while AY had a low return in 2017 and then a loss in 2018. 

### Refactoring
Code Module 1 (Original Code)

![code_module1](https://user-images.githubusercontent.com/85354946/156438915-afb17a70-a6be-4253-847e-3f110467e7ff.png)


Code Module 3 (Refactored Code)

![Code_refactored](https://user-images.githubusercontent.com/85354946/156438883-f81524fc-ca90-4cc6-aa34-7dd5241a0b4e.png)


### Execution Times
![![2017_refactored_time](https://user-images.githubusercontent.com/85354946/156414402-178f4189-4d36-4079-a83a-a066f9dd89c7.png)
![2018_refactored_time](https://user-images.githubusercontent.com/85354946/156414466-1d5064b5-b7c8-41f1-8718-8cbe5b978bed.png)
insert picture of original time 2017 ![2017_module1_code_run_time](https://user-images.githubusercontent.com/85354946/156437961-a0bc0eaa-fcf8-4758-924f-5a1e4e370d2f.png)

insert picture of original time 2018![2018_module1_code_time](https://user-images.githubusercontent.com/85354946/156437819-d1eb4cdd-09d2-4d1d-86a9-ad627792d7ec.png)


## Summary
### Advantages of Refactoring
If used on future projects with larger numbers of stocks it will be able to handle them better
code can be recreated more simply; easier to understand by those viewing it-- this applies to the assignment because it was easier to understand 
use less memory - run time was shorter
learning how the code works - ability to go through line by line and understand how the code works
### Disadvantages of Refactoring
Time it takes to refactor - spent days/ hours on this assignment
might create new errors or issues - debugging over and over. intially my outcomes did not match the original outcomes and i had to keep going back through and modifying
