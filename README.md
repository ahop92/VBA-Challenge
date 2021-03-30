# VBA-Challenge

## Background (added manually on 3/29/21)

The VBA script written sifts through the raw stock data provided in .xlsx format to summarize the progress of different stocks in a given year. The file was adjusted to be macro enabled in order to successfully 
run a VBA script.

Stock Data File: Multiple_year_stock_data.xlsm

### Code Functionality

The code summarizes the data via two new summary tables inserted into each worksheet with the following categories: 

Table 1: ticker, yearly change, percent change, total stock volume

    * The ticker is the associated stock for all of the summarized information 
    * Yearly change was calculated by subtracting the opening price on the opening day for the stock from the closing price on the last day recorded for that stock. 
    * Percent change was calculated by taking yearly change, dividing it by the aformentioned opening price and multiplying that ratio by 100. 
    * Total stock volume was calculated by summing all of the total stock volumes recorded for a given stock ticker 


Table 2: ticker, value, greatest % increase, greatest % decrease, and greatest total volume 

    * Greatest % increase is the greatest percent change seen in table 1 
    * Greatest % decrease is the most negative or smallest percent change found in table 1 
    * Greatest total volume is the greatest total stock volume found in table 1 
    * The assoicated stock ticker for each greatest metric is paired with the value


### Results 

Results of the code for 2014: 
![2014] (2014_VBA_StockResults.png)


Results of the code for 2015: 
![2015] (2015_VBA_StockResults.png)


Results of the code for 2016: 
1[2016] (2016_VBA_StockResults.png)
