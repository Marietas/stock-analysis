# Stock-Analysis - Challenge 2

## Overview of Project
This project aims to compare the stock performance between 2017 and 2018 to determine the best investment option. In terms of the code, the goal is to ensure that the script is more efficient than the original one by improving the logic of the code, making it easier to read, and reduce the execution time.  

## Results
To compare the stock performance between 2017 and 2018 and provide a refactored analysis, I worked with the original script called “challenge_starter_code.vbs” and renamed it to “VBA_Challenge. Vbs” afterward. 
To provide a more efficient code able to reduce the execution time, I used the function “tickerindex” that was applied for the following variables: “tickerVolumes”, “tickerStartingPrices,” and “tickerEndignPrices”. By using this function and applying the For loop and conditional analysis with [If then], I was able to structure and organize the original script to make it 4x faster. 

Please refer to the solved code shown below

![](https://github.com/Marietas/stock-analysis/blob/main/Resources/Script%20solution.PNG)

#### Comparison of the stock performance between 2017 and 2018
![](https://github.com/Marietas/stock-analysis/blob/main/Resources/Data%202017.PNG)

![](https://github.com/Marietas/stock-analysis/blob/main/Resources/Data%202018.PNG)

#### Execution times

##### 2017

![](https://github.com/Marietas/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

##### 2018

## Summary
### What are the advantages or disadvantages of refactoring code?
#### Advantages
- It makes the code easier to understand as the script becomes cleaner and more organized.
-	Helps to find bugs
-	Helps to improve the execution time
-	It makes the code more flexible to adapt to changes or updates of the software. 

#### Disadvantages:
-	Time-consuming when the code is large and complex
-	It is easy to misinterpret if the original script is not well explained and the developer cannot fully understand the procedure. 
-	The cost and time of refactoring can be higher than creating a new code.
-	Refactoring requires a substantial test procedure as well a time to complete the tests.

### How do these pros and cons apply to refactoring the original VBA script?
The biggest pro of refactoring the code was reducing the execution time by estimated four times compared to the original script. On the other side, the biggest con was to try to understand the original script's logic. In some instances, I believe that the time of refactoring could have been higher than creating a new code from Scratch. 
