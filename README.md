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

## Challenges
There are various challenges to consider when preparing this information; an example of a difficulty identified when accumulating this data includes breaking down categories that would be specific to Louises' needs. Given that Louise is interested in Kickstarter campaigns relevant to her play *Fever*, a lot of the data within the Kickstarter data set would not be entirely relevant to her. Thankfully, we were able to break this data down by sorting out Theater categories, and even more specific, by Plays.  

# Results

## Theater Outcomes by Launch Date
- Historically May has been the most effective start date for Theater Kickstarter campaigns.
- Historically December  has been the least effective start date for Theater Kickstarter campaigns.
- We conclude that summer months (May-July) have been the more effective start date for Theater Kickstarter campaigns.

## Outcomes based on Goals
- Theater Kickstarter campaigns are successful more frequently when the funding goal is less than $1,000.
- Theater Kickstarter campaigns historically become less effective as the goal increases to $29,000.

## Limitations
- Louise will not be able to determine specific region based effectiveness of Kickstarter campaigns, based on this dataset.
- This dataset does not provide data beyond 2017.

## Further Potential Analysis
Other helpful tables or graphs that we could generate with this data include:
- How many backers, on average, result in a more successful campaign.
- What average donation quantity results in a more successful campaign.
- We could also determine if a specific time of year results in higher backers & donation averages for Theater Kickstarter campaigns. 
