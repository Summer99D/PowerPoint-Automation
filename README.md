## instructions on how to use! 

Hello! this is a simple sample of automating powerpoint files. I have used S&P500 index and stock prices that cover dates from 2014-12-22 to 2024-12-20

**Warning**
For some reason, Apple stocks are not accesible. I am working on finding the issue. please try other tickers!

You can see the full list of companies [here](datasets/companies.csv).



you will have the chance to input a single date and get these in the powerpoint sildes:


 I. a chart of comparison between the desired stock price Vs. average stock price.

 II. a sentence mentioning the stock symbol(ticker), the date, and whether it has under or overperformed S&P500 average

 III. price of a certain stock across time (weekly, monthly, annualy) and a comparison agasint S&P index


 Instructions: 

 In order to generate a powerpoint file for a specific date and stock, please activate the virtual environment, install the
  [requirements]("/Users/samarnegahdar/Documents/school/PowerPoint-Automation/requirements.txt")
, and run this command in the terminal:

 **python automation_script.py --symbol XXX --date YYYY-MM-DD**




 ## Bibliography

 The datasets I have used are:

[S&P500-index](/Users/samarnegahdar/Documents/school/PowerPoint-Automation/datasets/sp500_index.csv)


[S&P500-stock](datasets/sp500_stocks.csv)



The merged data can be found at: 

[Merged data]("/Users/samarnegahdar/Documents/school/PowerPoint-Automation/datasets/merged_sp500.csv")



Later note: I converted .csv files to parquet to push into the repo






