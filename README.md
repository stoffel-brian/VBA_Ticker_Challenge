# VBA_Ticker_Challenge

## Overview of Project 
 - A summary of select stock ticker with total daily volume and return percentage.
 - The purpose of this project is to complete the requirements for Module 2. 
  
 [VBA_Challenge (.xlsm)](/../main/VBA_Challenge.xlsm)


## Results

The results of the projects are created for both 2017 and 2018. The VBA script loops through the rows and per ticket, sums the volums and calculates the return percentages. 

At a top level result analysis, 2017 had higher returns on average but a lower total daily volume in comparison to 2018.

[VBA Challenge 2017 (.PNG)](/../main/VBA_Challenge_2017.PNG)

[VBA Challenge 2018 (.PNG)](/../main/VBA_Challenge_2018.PNG)


## Summary

Refactoring the VBA script is intended to enhance the performance of the data summation and analysis. The disadvantage is the time cost to refactor the VBA script. The advantage is the time enhancement. The refactored VBA script includes for loops which tends to be faster. Specifically, for loops in VBA code allows blocks of code to run a set amount of times. The for loops work in this VBA script beacuse there is a known amount of tickers. An inherent disadvantage, if there was an unknown amount of tickers, then this VBA refactored script would not process. Another inherent disadvantage of VBA scripts is the difficulty of handing off the project to a colleague depends largely on shared VBA skills and notes of the logic/process within the script. 
