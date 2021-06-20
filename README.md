# stock-analysis
An analysis of ' green' stocks for Steve
## Overview of Project
  Steve really liked the workbook we prepared for him to analyze 'green' stock trends for 2017 & 2018. In this report, we refactored our subroutine so that Steve and his parents can continue to use this workbook to loop through a Stock Market Database.

## Purpose
Steve's parents want to invest in 'green' stocks. Our original macro allowed us to analyze if Steve's parents' choice to invest in DAQO green stocks was in their best interest. After determining that DQ was not the best stock, we refactored the code to analyze a larger amount of stocks, with the option to select and compare the stocks trend over years. This code will hopefully guide Steve and his parents into which 'green stocks' are best for them to invest in. We will also determing whether the refactored VBA code will run data faster, and if it is more efficient than the original VBA code.

Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Summary of Analysis

Our original analysis of the the DQ ticker's stock showed that DQ's stock dropped 63% in the past year, as shown in the phot below:

![DQ Ticker](https://user-images.githubusercontent.com/84881187/122668503-61b4e500-d186-11eb-95f8-27b7f88e9587.PNG)

Knowing this, Steve requested help identifying which 'green stocks' his parents should invest their money in. We refactored our original code to parse through the Stock Market Database, identify the Green Stocks, and total their volume. This would allow us to analyze the data, and identify the best stocks to invest in. See the tables below:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/84881187/122668899-6aa6b600-d188-11eb-99a2-1bee56b109ec.PNG)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/84881187/122668902-71cdc400-d188-11eb-8af7-1e7870d5db63.PNG)

Looking at our datasets, the best stocks for Steve's parents to invest in across 2017-2018 would be ENPH and RUN for their growth. We also examined how long it takes both the original macro, and the refactored macro to analyze the data:

Original vs Refactored Time for 2018 Data
![2018_Original Time](https://user-images.githubusercontent.com/84881187/122669627-45b44200-d18c-11eb-9e33-2889dfc71f3b.PNG)
![2018_Refactored_Time](https://user-images.githubusercontent.com/84881187/122669632-49e05f80-d18c-11eb-8ce8-1311f8233151.PNG)

Original vs. Refactored Time for 2017 Data
![2017_Original Time](https://user-images.githubusercontent.com/84881187/122669654-62507a00-d18c-11eb-8feb-610e4fbcd7b9.PNG)
![2017_Refactored_Time](https://user-images.githubusercontent.com/84881187/122669661-654b6a80-d18c-11eb-81b0-2f217c16ac76.PNG)


## Results

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

[the refactored time took marginally longer. More data to parse through, more memory used]



Some advantages and disadvantages of refactoring code are:

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

#### Advantages:

*Refactoring pre-existing code allows us to improve or update the code without chaning, or even improving the efficacy, of the sub-routine.
*By refactoring code, we can improve it's readability, making it easier to understand. If someone else needs to use our code in the future, or if we need to come back to it months later, having it refactored to be easy to undertand will allow us to make any needed changes with ease.
*Refactoring code allows us to find and fix any bugs. While a macro may be running correctly, there may be a line of code that could be refactored to more efficiently account for future variables.

#### Disadvantages:

*Refactoring can affect the intended outcomes of an analysis.
*Refacroting can be time consuming, and lead to creating more problems than they solve. While we may be aiming to have a code become more efficeint in analyzing a specific data trend, we may destabalize the macro's complex coding sets.


## Conclusion Summary: 

In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
