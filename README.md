# **Stock Analysis**
## **Project Overview:**
This project has two purposes:
1. Stock Analysis: Analyze 12 stocks based on their trading volume and average daily returns during 2017 and 2018.
1. Code Execution Performance: Refactor the underlying code to improve processing time.
___
## **Results:**
### **Stock Analysis**
As shown below, nearly all of the stocks performed much better in 2017 compared to 2018. While this analysis is limited to only two years, it appears these stocks react similarly within the stock market.  Diversification with bonds or stocks in different market segments may mitigate the negative impacts of stock market volatility on investment returns.  

**2017 Results:**

![Results2017](https://user-images.githubusercontent.com/90986041/134818363-21d51b71-0035-4d6b-80c5-81f188c53af6.png)

**2018 Results:**

![Results2018](https://user-images.githubusercontent.com/90986041/134818262-6be91942-e87e-4dd2-b357-ddd15579fc1f.png)

(See Refactored data source file link below.)

### **Code Execution Performance**

Below are examples of ways to improve code execution performance:
* *Minimize the number of times a worksheet is activated.*  I moved the "Worksheets("All Stocks Analysis").Activate" command line from the beginning of the script to the end into section (4) "Loop through...arrays to output...".  With this change, the data sheet was activated once and the output sheet was activated once.  The duplicative "Activate" command line in section (4) was not needed. I also moved the 'Create a header row into section (4), after the output worksheet was activated.
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
 * *Remove unnecessary For Next loops.*  Rather than create a separate For Next loop to initialize the tickerVolume to zero, I embedded the tickerVolume = 0 operation into the 'Ticker loop section after (2a).
   
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        ' Ticker loop
        For i = 0 To 11
            ticker = tickers(i)
            tickerVolume(tickerIndex) = 0
        
        ''2b) Loop over all the rows in the spreadsheet.
 * *Remove unnecessary If Then statements.*  Before section (3d) I did not add an If Then statement to determine if the next row's ticker didn't match.  It wasn't necessary because the For Next loop in section (2a) had already gone through all the rows.  I simply added tickerIndex = tickerIndex + 1 to increase the tickerIndex.

        Next r
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            ' An If Then is not necessary because the previous For Next loop has already gone
            ' through all the rows.  That is, at this point, it's the end of the rows.
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            'End If
        Next i
After making these small changes, the run time performance improved. 

**2017: Original Results - 1.28 seconds**
![Original_2017_Run_Time](https://user-images.githubusercontent.com/90986041/134818085-d32e0035-ee02-4329-99b3-d47b0f658d24.png)

**2017: Refactored Results - 0.82 seconds**
![VBA_Challenge_202017](https://user-images.githubusercontent.com/90986041/134818096-e895b63a-836c-4a2d-97a4-c605f278ca49.png)

**2018: Original Results - 1.22 seconds**
![Original_2018_Run_Time](https://user-images.githubusercontent.com/90986041/134818101-d640b0d8-d7b3-475f-9bde-087607e7d7ea.png)

**2018 Refactored Results - 0.81 seconds**
![VBA_Challenge_2018](https://user-images.githubusercontent.com/90986041/134818109-f2741aed-2bf2-4f9d-a27f-3fefcee2d09b.png)

While a 1/2 second improvement may not sound impressive, it is proof that being mindful of code order, iterations, and eliminating redundant operations will improve processing time.

___
## SUMMARY
Increasingly end-users are expecting instantaneous results when they click on a button. Therefore, coders must strive to write efficient scripts that effectively utilize system resources *and* generate results quickly.

Refactoring (editing) someone else's code set has its advantages and disadvantages, such as:

|Advantages | Disadvantages|
|:---:|:---:|
|Fresh set of eyes quickly find missed coding oppportunties |Change intended code (outcome) result
|Improve execution performance|Increase execution time|
|Opportunity to add clarity (notes, white space, code arrangement)|Create new bugs 

Before beginning the refactoring process, these are very important steps: 
1. Make a copy of the original script in case you need to revert back to it. 
1. Save a copy of the results so you can validate the outcome of the refactored code set is working as intended. 
1. Obtain original script run time(s) to ensure changes add efficiency. 

While refactoring the original VBA script, I experienced each disadvantage listed above. Some of my changes created bugs, due to new coder errors.  When I initially revised one of the For Next loops it stopped too soon and the last row of data didn't populate on the output sheet. After one of my first changes, while the results were correct, the run time was longer than the original script run time.  Without copies of the original script, results, and run times, I would have unintentionally made things worse.  In the end, I prevailed and achieved all of the above-mentioned refactoring advantages.
___
**Data Sources:** 

[Refactored File] https://github.com/SJLewer/stock-analysis/blob/main/VBA_Challenge.xlsm

[Original File] https://github.com/SJLewer/stock-analysis/blob/main/green_stocks.xlsm

**Analyst:** S. Lewer
