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

The main advantage of refactoring is efficiency. The efficiency can occur in the code itself or the readability of the coder. Refactoring code can help make the program run faster by changing up your algorithm. This is especially important when using larger sets of data. For example, if I can refactor my code to loop over the data set once as opposed to 10 times the code with execute much faster. Also, we can refactor code for the person/people who read the code. By taking out unnecessary code or adding spaces/comment, we can make organizing or reading code easier. This can avoid headaches attempting to figure out what the code is doing in the future.

Refactoring code can have it's drawbacks. The main disadvantage of refactoring code is time. The code already works, so why bother taking all the time to change it! There could be other projects with code that doesn't even run and thus it would be better to spend time on those. Also, we implementing any new code there is a chance that new bugs will occur. These new bugs could limit the functionality of the existing code you are trying to refactor. 

### How Pros and Cons Apply to Original Script

So how do the above pros and cons impact the vba script that I just refactored? For efficiency of the code, the analysis above shows in *Original 2017 and 2018* vs *Refactored 2017 and 2018* that the refactored code ran quicker. In fact it took about 9 times longer for the original code to run vs the refactored. On the down side, it took another 2 hours to complete. When I attempted to change all the code from the vba script I ran into multiple variables being named incorrectly causing several annoying bugs in my code. For a smaller data set of just 2017 and 2018 for 12 renewal energy stocks the original code worked just fine, however if we want to increase the number of stocks or the number of years, having the refactored code will save us plenty of time down the road.