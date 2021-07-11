# Stocks Analysis Project

## Overview of Project

Our client, Steve, is interested in the creation of an Excel workbook and VBA program that can analyze large amounts of stock data quickly and accurately. As such, this project aims to improve (also "refactor") existing code to better suit his goal.

### Purpose
The purpose of this analysis project is to improve upon previous code by reducing the run time needed to analyze the same data. 

### Results

- In general, the year 2018 seems to show poorer performance in green stocks compared to the year 2017. While "TERP" consistently yielded lower returns in both years, "ENPH" and "RUN" are stocks that continue to do well in 2018. Therefore, based on current data, "ENPH" and "RUN" seem to be the more reliable stocks out of the 12. 
- Execution times changed quite drastically after refactoring. In the previous code, the 2017 analysis took 1.71875 seconds and the 2018 analysis took 2.765625 seconds. The refactored code reduced the run time to 0.1523438 seconds for 2017 data and 0.1289062 seconds for 2018 data. The improved code allowed for the same function to be operated in approximately a tenth of the original time. 

<img width="415" alt="VBAChallenge_2017" src="https://user-images.githubusercontent.com/84816495/125179834-75210200-e1c0-11eb-9da1-2ea935354cda.png">
<img width="414" alt="VBAChallenge_2018" src="https://user-images.githubusercontent.com/84816495/125179837-7eaa6a00-e1c0-11eb-9779-24acfcb38ee0.png">


### Summary

#### Advantages and Disadvantages of Refactoring
1. The advantanges of refactoring code include: 
* shortening run time so analysis can be carried out in a quicker manner therefore allowing for larger and more complicated datasets to be analyzed.
* improving the logic progression of code and utilize more sophisticated and effective design to achieve the same result.
* the act of refactoring and analyzing one's own code can improve skills in VBA and general programming logic.

2. The disadvantages of refactoring code include: 
* in the process of changing the code, new inconsistencies or flaws in logic can be introduced which will require more time to troubleshoot.
* the refactoring process needs to be well-documented in order to track how and what changes were made. This ensures that revisiting the code at a later date will still allow the user to follow the development of the code.

#### How do these pros and cons apply to the original and refactored VBA code?
As seen in the results above for this project, refactoring the original VBA code did allow for the analysis for both 2017 and 2018 to be run quicker. This shows that if done correctly, refactoring can reduce run-time and allow for more efficient analysis of large amounts of data. The act of refactoring also allowed me to re-evaluate the logical reasoning behind my code and consider ways to improve. On the other hand, changing the code during refactoring did lead to minor errors that needed to be corrected. And coming back to the refactoring process after a day can lead me to lose the logic I was going for unless I appropriately indicated with comments my plan for the code.
