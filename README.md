# Stock Data Analysis - VBA Challenge

In this assignment, a VBA script is used to analyze quarterly stock data, apply conditional formatting, and generate summary statistics about stock performance.

## Instructions

This VBA macro processes stock market data over four quarters and outputs the following information:
- **Ticker Symbol**: Displays the stock's ticker.
- **Quarterly Change**: Calculates the change in stock price from the opening price at the start of the quarter to the closing price at the end of the quarter.
- **Percent Change**: Calculates the percentage change from the opening price to the closing price.
- **Total Stock Volume**: Sums up the total stock volume for each ticker.

The script loops through the data for each quarter, calculates the necessary values, and applies **conditional formatting**:
- Positive changes are highlighted in green.
- Negative changes are highlighted in red.

### Additional Functionality
The VBA script also identifies the following for each quarter:
- The stock with the **Greatest % Increase**.
- The stock with the **Greatest % Decrease**.
- The stock with the **Greatest Total Volume**.

The output is presented in a summary box next to the data.

### How to Use the VBA Script
1. Open the Excel workbook containing the stock data.
2. Run the VBA macro to analyze the data across all four quarters.
3. The results, including the calculated values and the highlighted cells, will be displayed on the sheet.

## Files in this Repository
- `Modulel1.bas`: Contains the VBA script that processes the stock data and applies conditional formatting to the stock data
- **Screenshots**: Includes screenshots of the Excel sheet with the calculated columns and conditional formatting applied.
- **README.md**: This file, providing an overview of the project.

## Summary of Results
- The stock with the **Greatest % Increase** is identified and displayed.
- The stock with the **Greatest % Decrease** is highlighted.
- The stock with the **Greatest Total Volume** is shown in scientific notation.

## Conclusion
This VBA script provides an efficient way to analyze stock market performance across multiple quarters, apply visual cues for quick interpretation, and summarize key statistics.

