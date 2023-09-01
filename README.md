# VBA-challenge
First initiate variables
Start Loop for every Worksheet in this Workbook
Variables take their values, identify the last not blind cell in worksheet
Starting Loop across the rows
For each of the row, assgning ticker value, accumulating a `total_volume` and checking if the next row(with ticker value) is not the same
  For passing this condition entering wrapping the Ticker values and reseting certian varaibles to default values:
    * calculating `price_change`
    * calculating `percentage`


    * Rendering a `TargetRow` with the data for the specific Ticker (ticker, price_change, percentage, total_value)
    * coloring the cell into red/green based on percentage value ( below 0 or above 0)
    * reseting ticker specific variables (total_volume = 0, next TargetRow, FirstTickerRow = CurrentRow) and incrementing general counter (Row)


Now we have filled out the table per ticker based data, we have to make another pass across this data to figure out greates max/min and ticker across the worksheets

Assigning default values to the greates(max/min) vars and tota value
in the current worksheet starting a loop from The previous table filled  (between Row and TargetRow which is last record written fro last Ticker)

 for every pass between ticker agreagated values checking and re-assigning min/max greatest and total_value into the vars assigned before the loop if the condition matches. re-printing the output into the table for each matched the condition.


That's it.
Thanks












