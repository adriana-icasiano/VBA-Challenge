# VBA-Challenge - instructions

* In this homework, students were giving stock data including the ticket symbol, daily opening and closing price and daily stock volume. Students were asked to create a script that will loop through all the stocks for one year and output the following summary table.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

  * Conditional formatting that will highlight positive change in green and negative change in red.

* Loop through new summary created above and output "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 


* Make the appropriate adjustments that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

# Summary of Codes used
 * For each ws in Worksheets, Next ws
 * lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
 * For loops

