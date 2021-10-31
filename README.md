# stock-analysis
Bootcamp Module 2
## Overview of Project: 
The project started by creating a workbook to analyze and format volume and retrun data for a select groups of twelve stocks over a given year. The purpose of this analysis was to refactor the original code for use with an expanded data set. We wanted to provide updated code that, when moving from twelve stocks to potentially thousands of stocks, would equal or outperform the original code in regards to analysis run time. We waned to see if the refactored code would prevent increased execution times and potential errors from running too long.
## Results: 
By selecting a pre-configured form-contol button, the analysis provided a formatted table showing Total Daily Volumes and Return percentages of the twelve selected stocks for the selected year.
![Workbook image](https://github.com/jmueller187/stock-analysis/blob/main/Resources/AllStocksAnalysisImage.png)

After supplying these results, we were also asked to determine if the runtime of the code could be reduced by refactoring the code in order to evaluate a larger sample of stocks.
The process started by evaluating the data for each of the twelve stocks over an entire year and determining run times of the initial code to set our benchmarks. Those results are as follows:

![Initial runtime for 2017](https://github.com/jmueller187/stock-analysis/blob/main/Resources/InitialTimeAnalysis2017.png)
![Initial runtime for 2018](https://github.com/jmueller187/stock-analysis/blob/main/Resources/InitialTimeAnalysis2018.png)

We then moved on to refactoring the code to evaluate the data for each year only once and capture data for each stock in order to update the summary table. The results of the refactored code are seen below:

![Refactored runtime for 2017](https://github.com/jmueller187/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Refactored runtime for 2018](https://github.com/jmueller187/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

These results show a 79% runtime reduction for 2017 data and 81% runtime reduction for 2018 data with the refactored code, while maintaining the functionality and output of the initial code.
The Percentage of runtine reduction was calculated by subtracting the refactored runtimes from the intial runtimes to obtain the overall decrease. Then each overall decrease was divided by the initial runtimes and muliplyed by 100 to obtain the final runtime reduction precentages.
## Amaysis Summary:
### What are the advantages or disadvantages of refactoring code?
The advantes of refactoring code include indreasing performace time for use with larger amounts of data while maintaining original design and functionality. Disadvantages include spending additional time and resources to refactor the original code, and potentially introducing errors/bugs to the original code during the refactoring process.
### How do these pros and cons apply to refactoring the original VBA script?
Refactoing of this particular code resulted in an average runtime reductions of 80% while still achieving the original indended results. The refactored code also ended up being easier to read by including additional variables for calculating output data and removing a nested for loop from within the initial code.