# Stock-Analysis - 2017 and 2018
## The purpose of this project is to show the volume of trades and the year end price vs the beginning close price percent change for a set of tickers representing 12 companies. Additionally, we are showing the performance of the original design vs refactored code, and discussing the advantages and disadvantages of refactoring.
## Results
### For 2017, we can see that the average volume was 263,886,592 for this set of stock tickers. All but one of the tickers had positive change during 2017. The average percent change at the end of 2017 vs the beginning was 67.32%. As the image below shows, the run time for the refactored code is far less than the original code.
### For 2018, we see a different story. The average volume for these tickers was relatively stable at 275,503,183. However, all but two stocks had negative percent changes when comparing the end of 2018 vs the beginning of 2018 stock prices. The average percent change for this set of stocks was -8.51%. Similar to 2017, the run time for the refactored code was far less than the original code. 
## Summary
### Advantages and disadvantages of refactoring code
#### Advantages
1. Refactoring can improve the design. 
2. Refactoring can make code more understandable and efficient. 
3. Refactoring can increase speed of execution 
4. Refactoring can change the way a developer thinks about coding. 
#### Disadvantages
1. Refactoring can be time consuming. 
2. Refactoring can likely clutter some code that would be easier to understand if a portion was set in a different subroutine or module. 
3. As a result, it can be more difficult to debug code or pass on to another developer for maintainence. 
### Advantages and disadvantages of refactoring VBA scripts. 
#### Advantages
1. Refactoring increases the performance of excel, and makes it more useful. 
2. Refactoring VBA can mean the need for less buttons within the excel spreadsheet, making the use of the subroutines easier. 
3. Refactoring makes it easier to pass models once turnover occurs. 
4. Code can be made more efficient by reducing things like how many times VBA cycles through code. 
#### Disadvantages
1. It can add to confusion by putting too much into a macro, making it less easy to follow and debug. 
2. It likely increases the time required to debug. 
3. Because a macro may be efficiently accomplishing many tasks, it may make it more difficult to re-use code in other modules. 
