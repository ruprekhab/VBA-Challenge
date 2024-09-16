# VBA-Challenge
VBA script for stock Analysis
Excel Stock Analysis VBA Script

This VBA macro analyzes stock data and generates a summary report for each stock ticker. It calculates and displays Quarterly price change, percentage change, total stock volume for each Ticker. It highlights positive percentage change in green and negative change in red. A separate table gets created highlighting tickers with the greatest percentage increase, decrease, and the greatest total volume. 

How It Works

The script loops through all worksheets in the workbook.

For each ticker symbol, it calculates:
1) Total stock volume traded.
2) The price change between the opening price at the beginning of quarter and closing price at the end of quarter.
3) The percentage change between the opening and closing prices.

A summary is created for each ticker in the Summary table, which includes:
1) Ticker symbol.
2) Price change.
3) Percentage change.
4) Total volume traded.

Conditional formatting is applied to the percentage change column, highlighting positive changes in green and negative changes in red.

The script also finds and displays:
1) The ticker with the greatest percentage increase.
2) The ticker with the greatest percentage decrease.
3) The ticker with the highest total stock volume.
