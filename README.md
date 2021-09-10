# Stock Analysis

## Overview of Project: 
The purpose of this project consistis of two parts:

The first one is to analyze the performance of the stocks listed in our Macro in the years of 2017 and 2018. 
The second goal is to refactor the original code in the Macro to determine what is the effect of vectorizing the variables used to read through each row and to calculate the Volume and the Variation of each stock.


### Results on stock performance and code execution times.

#### Stocks performance

By running the macro in the Macro-Enabled Excel file, we'll get to these results:

All Stocks Analysis - 2017      |  All Stocks Analysis - 2018
:-------------------------:|:-------------------------:
![(All Stocks Analysis - 2017)](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2017%20Analysis.png)  |  ![(All Stocks Analysis - 2018)](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2018%20Analysis.png)

Calculating the average on of the total daily volume and the return for all the stocks in 2017 and 2018 we come to two main conclusions:

* Overall, the stocks performed much better in 2017, where most of them had positive returns.
* On average, the return of all the stocks in 2017 was 67% with a daily volume of 263,886,592 while in 2018 the return came down to -9% with a volume of 275,503,183. That means the return were much higher in 2017 with around 95% of the volume in 2018.

While isolating the DQ stock:
* In 2017, the DQ stock provided 199.4% of return with a volume of 35,796,200.
* In 2018, the DQ stock had -62.6% of return with a volume of 107,873,900. That means there were much more volume of transactions, but the return was much worse in 2018, when compared to 2017.

#### Code execeution time.

The execution times of the All Stocks Analysis Macro before the refactoring can be seen below: 

Original code - 2017       |  Original code - 2018
:-------------------------:|:-------------------------:
![2017 All Stocks Analysis Macro Performance](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2017%20Analysis%20Timer.png)  |  ![2018 All Stocks Analysis Macro Performance](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2018%20Analysis%20Timer.png)

And here are the results after the refactoring, for 2017 and 2018, respectively:

Refactored code - 2017       |  Refactored code - 2018
:-------------------------:|:-------------------------:
![2017 All Stocks Analysis Refactored Macro Performance](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2017%20Refactored%20Analysis%20Timer.png)  |  ![2018 All Stocks Analysis Refactored Macro Performance](https://github.com/ericosabino/stock-analysis/blob/main/Resources/2018%20Refactored%20Analysis%20Timer.png)

We can see that the refactored code executed ~5 times faster in 2017 and ~7 faster in 2018, so a much better performance after the refactoring. The vectorized approach in the refactored code proved to be much more effective than storing the values in individual variables and iterating through them.

### Code Refactoring Pros and Cons

#### General Refactoring

Refactoring a code usually improves response, loading times, execution times, and other variables that significantly improve the internal performance of the software, resulting in a better user experience and, in the case of the Stock Analysis file, processing more information in less time.

The downside of refactoring is that usually it takes time and effort to create advanced code, so if the business have a tight deadline or budget/resources constraints, it may be challenging to refactor code, particularly when the code is already being used by a large user base.

When a new product is being launched, usually the sooner the Minimum Viable Product is available in the market, the sooner the company can get revenue or investments to run the business, so usually refactoring code is the last option the software team will focus on.

#### Regarding the original and refactored VBA Script

The biggest advantage of the refactored VBA Macro was to process more data in less time. We could see that the execution times were 5 to 7 times lesser in the refactored code, respectively, for 2017 and 2018. 

The main disadvantage in the refactoring process was that the code was much more complex than the original one, so it took much more time to understand the indexes, set up the arrays, and calculate the volume and return for each stock. Even though the code was shorter, it was much more complex.
