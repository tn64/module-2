<!-- Photo below by Lorenzo from Pexels -->
<img src=https://github.com/tn64/stock-analysis/blob/main/Resources/pexels-pixabay-210607-edit.jpg>

# DA Bootcamp Module 2: Stock analysis.

## Overview of Project
The original assignment was to use VBA script to determine the annual volume of shares sold and annual rate of return for a green energy stock. Then, the code was expanded to return the information for a list of green energy stocks. The script initially created to return the information for the list went through the entire list line-by-line. This was effective, but the question for the Module 2 Challenge was whether refactoring the code to loop through all of the data one time would make the code run faster, making it more useful for searching through a very large list of stocks.

## Results

The original script first created a timer, then initialized an array of green stocks, determined the number of rows in the data sheets, and used a nested for loop to go through each row line-by-line. It would search for the indicated ticker symbols, then the stock's volume, and then calculate the annual return for each stock. After each instance of the inner for loop the code would then ouput the ticker, annual total volume, and annual return of the stock. Finally, the script displayed a message box showing how much time it took to run the code.

### The Original Nested For Loop

        For i = 0 To 11
           ticker = tickers(i)

           Worksheets(yearValue).Activate
           For j = 2 To RowCount

            'get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
           End If

               'get starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                     startingPrice = Cells(j, 6).Value
                     End If

                 'get ending price for current ticker
                     If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                     endingPrice = Cells(j, 6).Value
                     End If

            Next j

         'output data for current ticker
         Worksheets("All Stocks Analysis").Activate
         Cells(4 + i, 1).Value = ticker
         Cells(4 + i, 2).Value = totalVolume
         Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

          Next i
      
Time to run this code for 2017:

<img src=https://github.com/tn64/stock-analysis/blob/main/Resources/2017_before_refactor.png>

Time to run this code for 2018:

<img src=https://github.com/tn64/stock-analysis/blob/main/Resources/2018_before_refactor.png>

### The Refactored Code

The challenge refactored the code by creating a ticker index (tickerIndex) and three output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices):

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Then, rather than using a nested for loop to run through each line of code, the refactored code first created a for loop to initialize the tickerVolumes to zero:

          For i = 0 To 11
              tickerVolumes(i) = 0
              tickerStartingPrices(i) = 0
              tickerEndingPrices(i) = 0
          Next i
          
Next a second for loop was created to take advantage of the ticker index to first find the volume, then find starting and ending prices (for the return calculation). and finally advance to the next ticker in the ticker index:

            For i = 2 To RowCount
    
                    'increases the current tickerVolumes                        
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                    
                    'check if the current row is the first row with the selected tickerIndex, then assign the current starting price to the tickerStartingPrices variable
                    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    End If
            
                    'check if the current row is the last row with the selected tickerIndex, then assign the current closing price to thetickerEndingPrices variable
                    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                    End If
                        
                    'increase the tickerIndex if the next row's ticker doesn't match the previous row's ticker
                     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                        tickerIndex = tickerIndex + 1
                    End If
                         
            Next i

Then, a final for loop was created to loop through the arrays to output the "Ticker," "Total Daily Volume, and "Return" for each of the selected stocks.

           For i = 0 To 11
              tickerVolumes(i) = 0
              tickerStartingPrices(i) = 0
              tickerEndingPrices(i) = 0
          Next i

The refactored code ran significantly faster. Time to run this code for 2017:

<img src=https://github.com/tn64/stock-analysis/blob/main/Resources/2017%20Refactored.png>

Time to run this code for 2018:

<img src=https://github.com/tn64/stock-analysis/blob/main/Resources/2018%20Refactored.png>

## Summary

### What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code may include:

- Making the code easier to understand and therefore making debugging easier
- Making better use of code patterns
- Making the code run faster
- Reducing code size

The disadvantage of refactoring code may include:

- Refactoring code can be time consuming
- Because it is time consuming, it may fail a cost-benefits analysis

### How do these pros and cons apply to refactoring the original VBA script?

After refactoring the code, it was easier to read and made more efficient use of code patterns. Though cost was not an issue, refactoring the original code for the exercise did take significant time. This is primarily due to unfamiliarity with VBA script. However, the time necessary to a) reinvision the scritp and b) determine how to write the new script would probably be significant even with greater familiarity with VBA script.
