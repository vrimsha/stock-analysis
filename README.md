# VBA Challenge
## Overview of Project
### Purpose: 
<br>The client wants to invest in Daqo's stock. Before investing in it they would like to expand the dataset to include the entire stock market over 2017 and 2018 years. All Stocks Analysis VBA script needs to be refactored.</br>

## Analysis 
<br> I have measured “All Stocks Analysis” original script performances for two years before refactoring it.</br>
<br>The script performance ran 0.59375 seconds for the 2017 year:</br>

![2017_result_before_refactoring](2017_result_before.png)

<br>The script performance ran 0.609375 seconds for the 2018 year:</br>

![2018_result_before_refactoring](2018_result_before.png)

<br>I followed a prepared code structure from VBA_Challenge.vbs.</br>
<br>Next step was refactoring the VBA code in order to make the new code more efficient (fewer steps, less memory, easier to read).</br>
<br>Please see below the refactored loops:</br>

![Refactored_loop.png](Refactored_loop.png)

<br>I have measured new performance results. </br>
<br>Refactored script performance ran 0.140625 seconds for the 2017 year:</br>

![VBA_Challenge_2017](VBA_Challenge_2017.png)

<br>Refactored script performance ran 0.140625 seconds for the 2018 year:</br>

![VBA_Challenge_2018](VBA_Challenge_2018.png)

## Results
<br>The refactored code runs faster by 0.23 seconds in 2017 and 2018.</br>
<br>I compared All Stocks “Return” results in 2017 and 2018.</br>

<br>In 2017 almost all 12 tickers went up in return by minimum 5%, maximum 199.4% except ticker: TERP. It went down by 7.2%.</br>   
 
<br>In 2018 you can see a different situation, when most of tickers had negative return except two tickers ENPH and RUN.</br>
 
<br>Daqo Stock had a high return 199.4% with 35,796.200 total daily volume in 2017. However, in 2018 it had -62.6% in return with 107,873.900 total daily volume. As a result, it is not best to time for the client to invest in the stock when return result is negative.</br>

## Summary
<br>Advantages of refactoring a code are the faster performance, clear code that easy to read, and an opportunity to work properly with thousands of stocks.</br>
<br>Disadvantages of refactoring a code are cost, risks and working with bugs.</br>
<br>The benefits of refactored VBA scrip are:</br>
-	The speed of performance is faster by 0.23 seconds. 
-	The code is easier to understand for future users
-	Helps to find bags
-	Find new ways to write a code

<br>The weaknesses of refactored VBA script are:</br>
- Fix bugs
-	Extra time 
-	Extra cost






