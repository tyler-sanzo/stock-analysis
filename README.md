# Annual Stock Analysis
Creating a tool using VBA scripting to analyze annual stock data.

## Project Overview
Steve, a friend who recently graduated with his finance degree, wants to help his parents make smart investment decisions within the green energy sector. Steve has provided us with a dataset of stocks and relevant financial info and has tasked us with creating a analytical tool which will be used to help his parents invest wisely.

## Results
Although this tool has been configured to be flexible and take more annual datasets, below I will go over the results of the analysis for the datasets given by Steve (2017 and 2018).

### 2017
2017 was a strong year for green energy stocks all around with all but 1 of the 12 companies having a positive return at the end of the year; four of which saw a return of over 100%. 

![2017 Results](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/VBA_Challenge_2017_Results.PNG)

Below are images taken of the analysis runtimes for the both the initial and refactored code, respectively.

![2017 Runtime Initial](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/initial_runtime_2017.PNG)

![2017 Runtime Refactored](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/VBA_Challenge_2017.png)

From the images we can see the refactored code increased run efficiency by close to 80%.

### 2018
2018 was a major change from the previous year showing 10 of the 12 stocks having a negative return. The two stocks which experienced positive returns, ENPH and RUN, both showed strong returns of greater than 80%. One thing to note about these two is that while their 2018 returns were similar, ENPH's YOY return decreased from 130% while RUN saw a healthy increase compared to their prior year's 5.5% return. 

![2018 Results](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/VBA_Challenge_2018_Results.PNG)

Again we see a similar runtime decrease using the refactored code.

![2018 Runtime Initial](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/initial_runtime_2018.PNG)

![2018 Runtime Refactored](https://github.com/tyler-sanzo/stock-analysis/blob/main/Challenge/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and disadvantages of code refactoring
In a general sense, refactoring code depends largely on the scope of the project and deadlines. In the real world, often times functionality is added in a working state but suffers from runtime efficiency and readability due to release deadlines. These issues are solved through refactoring; restructuring dirty code without changing the functionality/front end of the program. Refactoring definitely makes sense for projects with a large scope and deadlines which cause the dirty code to be released in the first place. On the other hand, it may not be worth the time and effort spent for the miniscule gain in efficiency as it pertains to coding projects with a smaller scope. 

### Refactoring the stock analysis VBA script
Refactoring as it pertains to the analysis in this project was largely done to test the efficiency of the original code. The main differences between the two scripts is the use of arrays in the refactored code to store values locally before outputting the array values.

#### Pros
The advantages of this can be seen in both years of the results section; after refactoring the code we see an 80% decrease in runtime. Although the actual time value difference is only a fraction of a second and can be considered negligible if this is were the extent of the project, this efficiency increase would be very beneficial if the dataset was much larger. 

#### Cons
The main disadvantage I can see in the refactored code is the use of "magic numbers" when setting up the arrays. If one wanted to use this script to analyze a dataset with more than 12 stocks, the code would have to be refactored again to accommodate for this. On the contrary, the original script is able to handle any number of stocks as one would like to add given they are added to the "tickers()" array. 



