# **Stock Analysis**
## **Project Overview:**
1. End-user perspective: Anlayze 12 stocks based on their trading volume and average daily returns during 2017 and 2018.
1. Coding perspective: Refactor the underlying code to improve processing time.
___
## **Results:**
### **Stock Analysis**
As shown below, nearly all of the stockes performed much better in 2017 than in 2018. 
 
 insert image files here

### **Code Execution Performance**

Listed below are examples of ways to improve code execution performance:
* *Minimize the number of times the worksheet is activated.*  I moved the "Worksheets("All Stocks Analysis").Activate" command line from the beginning of the subroutine to the end into section (4) "Loop through...arrays to output...". This meant the duplicative "Activate" command line in section (4) was not needed.  I also moved the 'Create a header row into section (4).
    ``` 
        Worksheets("All Stocks Analysis").Activate
        
        'Format the output sheet on All Stocks Analysis worksheet
        ' Worksheets("All Stocks Analysis").Activate - this row is not needed
        
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
        'Reset tickerIndex
        tickerIndex = 0
        
        'Populate ticker data into output sheet
        dataRowStart = 4
        dataRowEnd = 15
        
        For r = dataRowStart To dataRowEnd
            Cells(r, 1).Value = tickers(tickerIndex)
            Cells(r, 2).Value = tickerVolume(tickerIndex)
            Cells(r, 3).Value = tickerEndingPrice(tickerIndex) / tickerStartingPrice(tickerIndex) - 1
            tickerIndex = tickerIndex + 1
        Next r 
 * *Remove unnecessary For Next loops.*  Rather than create a separate For Next loop to initiatize the tickerVolume to zero, I embedded the tickerVolume = 0 operation into the 'Ticker loop section after (2a).
   
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        ' Ticker loop
        For i = 0 To 11
            ticker = tickers(i)
            tickerVolume(tickerIndex) = 0
        
        ''2b) Loop over all the rows in the spreadsheet.
 * *Remove unnecessary If Then statements.*  After the section (2a) For Next loop over all the row in the data worksheet, before section (3d) I did not add an If Then statement to determine if the next row's ticker didn't match.  It wasn't necessary because the For Next loop had already gone through all the rows.  I simply added tickerIndex = tickerIndex + 1 to increase the tickerIndex.

        Next r
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            ' An If Then is not necessary because the previous For Next loop has already gone
            ' through all the rows.  That is, at this point, it's the end of the rows.
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            'End If
        Next i
After making these small changes, the run time performance improved. 

*insert before and after image files here*

While a 1/2 second improvement may not sound impressive, it is proof that being mindful of code order, iterations, and eliminating redudant operations will improve processing time.

___
## SUMMARY
Increasingly end-users are expecting instanteous results when they click on a button. Therefore, coders must strive to write efficient code sets that meet (or exceed) end-user performance expectations *and* effectively utilize system resources.

Refactoring (editing) someone else's code set has its advantages and disadvantages:

|Advantages | Disadvantages|
|:---:|:---:|
|Fresh set of eyes quickly find missed coding oppportunties |Change intended code (outcome) result
|Improve execution performance|Increase execution time|
|Opportunity to add clarity (notes, white space, code arrangement)|Create new bugs 

Before beginning the refactoring process, there are very important steps: 
1. Make a copy of the original script in case you need to revert back to it. 
1. Save a copy of the results so you can validate the outcome of the refactored code set is working as intended. 
1. Obtain original script run time(s) to ensure changes add efficiency. 

While refactoring the original VBA script, I experienced each disadvantage noted above. Some of my changes created bugs, due to new coder errors.  When I initially revised one of the For Next loops it stopped too soon and the last row of data didn't populate on the output sheet. After one of my first changes, while the results were correct, the run time was longer than than the original script run time.  Without the copies of the original script, results, and run times, I would have unintentionally made things worse.  In the end, I prevailed and achieved all of the above-mentioned refactoring advantages.
___
Data Source: insert link to file here

Analyst: S. Lewer