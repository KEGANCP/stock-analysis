# Stocks Analysis
## Project Overview
Steve has for assistance preparing data analysis for his parents, who are interested in investing in green energy. After completing his college degree in finanace, his parents wanted to be his first clients. For various reason, his parents were interested in investing in a solar panel company, DAQO. In order to ensure this was a viable and lucrative option for his parents, Steve requested help to analysize stock data over the past couple years. Upon providing our initial findings, Steve requested to dig a little deeper to determine an overall analysis to include all stocks within our dataset. We will be utilizing VBA to improve our initial findings, to improve effectivness and provide more relevant details. 
# Project Objective
Utilizing [green stock dataset](https://github.com/KEGANCP/stock-analysis/blob/main/VBA_Challenge.xlsm), totaling in raw data for 12 green energy companies, we were able to quantify and analyze the effectiveness and return history of these companies to determine if their may be better investment options other than DAQO (DQ). Our initial file was effective for the smaller set of data, but there was oppurtinity to improve and refactor our VBA code for a more effective run time to provide details for all tickers. 
# Project Results
By enhancing our intial VBA code, I was able to reduce our run time to populate this data for our now larger set of green energy tickers. With this improved run time, Steve can now generate and quantify Total Daily Volume and Return percentage for each specific ticker in half the run time. In the below screenshot you will see our runtime timer results compared from our first VBA code to our enhanced code. 

![This is an image](https://github.com/KEGANCP/stock-analysis/blob/main/Resources/2017_Compare.png)

![This is an image](https://github.com/KEGANCP/stock-analysis/blob/main/Resources/2018_Compare..png)


## Refactored Code
The below will provide a visual of some of the refactored code which resulted in higher effeciency. This example specifically outlines the improved nested loop function.

![This is an image](https://github.com/KEGANCP/stock-analysis/blob/main/Resources/Refactored_Code.png)

# Project Summary
There are various elements to consider when preparing this information; 
### Advantages
- When refactoring, you're goal is to review the code in it's entirety; this allows an oppurtinituy to identify redundancies, or code that isn't neccessary to achieve the results your looking for. This also gives the oppurtinity to identify code that can be improved/consolidated to improve run time. When working with a large dataset, refactoring can be extremely effective and important to the end user.
### Disadvantages
- Depending on the size of your project, refactoring can be very time consuming, and may not always be worth the time if you're only able to improve run time by a fraction of a second. It's important to consider which codes you're looking to improve, and how to restructure appropriately as you could easily mis-place or mis type a code that could result in oulling incorrect or incomplete data.

### Recap of above
- Considering that Steve wants to utilize this tool for a larger set of data, I believe refactoring this code was important to improve effectiveness. While refactoring, I was able to identify a few redundancies and also an oppurtinity to improve & consolidate various code. As more data is added to this tool, these changes were vital to provide the data in a viable run time.
