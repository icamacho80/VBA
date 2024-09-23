Summary: the excel file has information for shares prices during four quarters. It includes a Macro in VBA code that sumarizes the performance of the ticket 
during the each quarter; its change in price in dollars and percentage between the opening price and the closing price at the end of the quarter. In addition, 
it creates a summary of the top one highest increase of price in percentage, highest decrease of the price in percentage and the highest volume.

    The steps of the macro include
    - Set dimensions, definitions of the loop variables, start, count of rows to automate the loop, and measures including percent and dailychage. As the macro
      runs into all pages of the workbook, we define the worksheets.
    - Runs a loop for in order to run the macro in each of the worksheets
    - Code to set the titles of the columns
    - Code to set the inital values for the loop, including start the counter in zero
    - Gets the row number in the last row with data. Uses formula countA which returns the numbers of rows
    - For all raws not empty, review if still the same value as the previous one.
    - with if evaluates of the cell has the same name ticket as the previous one 
    - if so then include to total value the cell (i,7) that holds the volume.
    - then printing in Colums I,J,K,L all the results that are being stored line per line 
    - if the cell is not the same as previous one then start the counter
    - Then calculate the changes in dolars and percentage of the share ticket negotiated
    - Then print the results accumulated in IJKL columns
    - set up green cells for positive numbers and red for negative changes in percentage 
     - Reset variables for new stock ticker
     - if ticker is still the same add results to the counter
     - go to next iteration
     - For Range Q2 = max percentage of change
     - For Q3 = lowest percentage of change
     - For Q4 = higest volume traded
     - Fixes format of last column
