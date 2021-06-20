# stock-analysis
An analysis of ' green' stocks.

## Overview of Project
  Steve really liked the workbook we prepared for him to analyze 'green' stock trends for 2017 & 2018. In this report, we refactored our subroutine so that Steve and his parents can continue to use this workbook to loop through a Stock Market Database.

## Purpose
Steve's parents want to invest in 'green' stocks. Our original macro allowed us to analyze if Steve's parents' choice to invest in DAQO green stocks was in their best interest. After determining that DQ was not the best stock, we refactored the code to analyze a larger amount of stocks, with the option to select and compare the stocks trend over years. This code will hopefully guide Steve and his parents into which 'green stocks' are best for them to invest in. We will also determine whether the refactored VBA code is more efficient, using less memory, and makes the data more readable for the user.

## Results of Analysis

Our original analysis of the the DQ ticker's stock showed that DQ's stock dropped 63% in the past year, as shown in the phot below:

![DQ Ticker](https://user-images.githubusercontent.com/84881187/122668503-61b4e500-d186-11eb-95f8-27b7f88e9587.PNG)

Knowing this, Steve requested help identifying which 'green stocks' his parents should invest their money in. We refactored our original code to parse through the Stock Market Database, identify the Green Stocks, and total their volume. See below for the code designed for the user to select which year's stock data they wish to analyze:

![Refactored_Code_YearCycle](https://user-images.githubusercontent.com/84881187/122672006-5a4a0780-d197-11eb-9549-023b13557a48.PNG)


This would allow us to analyze the data, and identify the best stocks to invest in. We also included buttons so that when using the macros, one can:

  1. Run/analyze the data witha button click, that will prompt the user to select which years data they would like to analyze.
  2. Format the data in the event a bug occurs within the macro, to increase the table's readability.
  3. Clear the table before analyzing another set of data.


**[See the tables with buttons below:]**


![VBA_Challenge_2017](https://user-images.githubusercontent.com/84881187/122672637-b2363d80-d19a-11eb-9376-783da10eb0cc.PNG)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/84881187/122672641-b95d4b80-d19a-11eb-894c-cdbd5f072bc6.PNG)


**[See below for the refactored code lines that analyze our data:]**

![Refactored_Code_ticker](https://user-images.githubusercontent.com/84881187/122671978-34bcfe00-d197-11eb-87f3-675785b4feaf.PNG)



Looking at our datasets, the best stocks for Steve's parents to invest in across 2017-2018 would be ENPH and RUN for their growth. Our refactored macro enabled us to analyze multiple tickers and select the year we wish to analyze. Instead of going into the code to determine which year and ticker we wish to analyze, our VBA code now provides a Message Box that will allows us to enter which year's data to analyze. The buttons provided on the workbook streamline the process for Steve and his parents. The refactored code is implemented so that they can use it go examine future trends as they collect and input more Stock Market Data. We examined how long it takes both the original macro, and the refactored macro to analyze the data:


#### **[Original vs Refactored Time for 2018 Data]**

![2018_Original Time](https://user-images.githubusercontent.com/84881187/122669627-45b44200-d18c-11eb-9e33-2889dfc71f3b.PNG)
![2018_Refactored_Time](https://user-images.githubusercontent.com/84881187/122669632-49e05f80-d18c-11eb-8ce8-1311f8233151.PNG)


#### **[Original vs. Refactored Time for 2017 Data]**

![2017_Original Time](https://user-images.githubusercontent.com/84881187/122669654-62507a00-d18c-11eb-8feb-610e4fbcd7b9.PNG)
![2017_Refactored_Time](https://user-images.githubusercontent.com/84881187/122669661-654b6a80-d18c-11eb-81b0-2f217c16ac76.PNG)

#### **[See below for code line to initiate the MEssage Box for time it takes to analyze data]**

![Refactored_Code_Timer](https://user-images.githubusercontent.com/84881187/122672433-7c448980-d199-11eb-8285-b989a7e3c93e.PNG)


Looking at the times, the refactored code took a little bit longer to analyze than the original data. This is because now our macro is analyzing data across all of the Stock Market Database tickers, as opposed to just data for DQ. This means there are more variables and more memory is needed to process our analysis. We can conclude that as more data, across more years are added, the datasets will take longer to analyze and populate.


## Summary

Our original analysis was able to inform Steve and his family that DQ was not the best 'green stock' for them to invest in. By refactoring our Stock Market Database VBA Macro, we were able to have our code loop through all of the tickers and generate the total volumes. This gave Steve and his parents a comprehensive table of the total volumes of desireable stocks. This table gave a guide on how stocks moved between 2017-2018, and where they should invest. Our refactored code also allows for Steve, or other future users of the code to be able to introduce more datasets of data for future years of Stock Market values. 

While refacorting our VBA code greatly helped Steve and his parents, some advantages and disadvantages of refactoring code are:

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
