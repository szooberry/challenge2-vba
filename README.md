# challenge2-vba

The VBA script in this repository analyzes yearly stock market data

If working properly, the script should yield the following analysis:
1. The unique ticker symbols associated with the stocks found in the raw data
2. The yearly difference in stock price for each stock
3. The percent difference compared to the opening price for each stock
4. The total volume traded over the year for each stock

The script will also yield the following additional calculations:
1. The stock with the greatest percent increase
2. The stock with the greatest percent decrease
3. The stock with the greatest trade volume for the given year

IMPORTANT CONSIDERATIONS

For the script to work properly, the data in each sheet must be formatted in the following way:

1. The first column must contain the ticker symbol of the stock and must be repeated for each data entry for the associated stock
2. The second column must contain the specific date associated with the stock price on that date
3. The third column must contain the opening price of the stock on the given date
4. The sixth column must contain the closing price of the stock on the given date
5. The stock's prices must be ordered chronologically by date, going from the beginniing of Jan to the end of Dec for the given year
6. Each sheet in the workbook must contain only the data for a specific year

Examples of what the final product should look like are provided in the PNG screenshots in this repository
