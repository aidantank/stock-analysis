# Green Stock Analysis

## Overview of Project

### Purpose

We have been tasked with analyzing green stocks for Steve. Using the data provided [here](/VBA_Challenge.xlsm), we will compare stock performance between 2017 and 2018. In the future Steve wants to expand the dataset to include more stocks and years. Due to the increase in data, we need to refactor our original vba script in the hopes of reducing the run time. The second part of our analysis will compare the run time of our original script and our refactored script.

## Results

### Comparing Stock Performance between 2017 and 2018

Below in *Table 2017* and *Table 2018* are condensed data for stocks in 2017 and 2018 respectively. The ticker head refers to the company, the total daily volume is the total number of stocks traded throughout the year, and the return is the change in stock value from the start of the year to the end of the year. The return is calculated using `Starting Price/Ending Price - 1`. Finally, the tables were formatted so positive returns are colored green and negative returns are red. We will be using the total daily volume and returns to compare stocks in 2017 vs 2018.

![](/Resources/AllStocks2017.PNG)![](/Resources/AllStocks2018.PNG)
> Table 2017 and 2018

When the colors are added to the tables it becomes very clear that the 12 green stocks as a whole performed considerably better in 2017 than 2018. 11 out of the 12 green stocks in 2017 had positive returns whereas only 2 out of 12 did in 2018. It would have been much more profitable to have a green stock in 2017 than in 2018.

### Comparing Execution Times of Original Script vs Refactored Script

What if Steve wants to know green stock performance in other years? Well we need to refactor our code so that is more efficient at processing larger amounts of data. Our original script had the bulk of the code in the format:

```
'Loop through tickers
For i = 0 To 11
    'Loop through Data
    For j = startData to endData
        'Perform Analysis
    Next j
    'Output data
Next i
```

I noticed here that we are looping through tickers and then through the data. Creating nested For loops increase the duration that the code processes exponentially. Although this works, it can be done more effieciently. Below is the duration the code took in the original script:

![](/Resources/OriginalScript2017.PNG)![](/Resources/OriginalScript2018.PNG)
> Original 2017 and 2018

The method I used to refactor the above code was removing the outer For loop in favor of an array. The code looked like this:

```
'Initialize Array of tickers and Index to track
For i = startData to endData
    'Grab ticker based on index. Starts at 0
    'Perform Analysis
    'Increment index to next ticker
Next i
'Output data from Array
```
Using an array you are able to grab different ticker information without having to loop through the data more than once. In the original script you had to loop through the entire data 12 times for the number of tickers! In fact, it took about 9 times longer to run the original script vs the refactored script if you look at the times below:

![](/Resources/VBA_Challenge_2017.PNG)![](/Resources/VBA_Challenge_2018.PNG)
> Refactored 2017 and 2018

However, both times are under a second. So with the data for a year only including 12 companies, it does not make too big of a difference. But if we mulitply the companies by 10 or take stock values by the hour instead of by the day, we will see a much larger difference in time to run the original script vs the refactored script.

## Summary

### Advantages and Disadvantages of Refactoring Code

The main advantage of refactoring is efficency and the main disadvantage is time. 

### How Pros and Cons Apply to Original Script

Approached the code from a different way, took less time to run. Took a lot more of my time to refactor the code and there only was a 0.8 second difference in runtime.

