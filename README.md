# Stock Analysis Project
## Overview
The purpose of this project was two-fold: 
- To perform an analysis of stock values for Steve, the client, to gauge performance for potential investment opportunities for his parents. 
- To create a report template so that Steve could then generate the reports himself. 

The details about the analysis, the report generator and the project code and related issues are documented in the next sections of this summarry.
## Stock Performance Analysis
Steve initially requested stock performance details for Daqo (ticker index DQ) because his parents were interested in investing in a green energy company. The results of the analysis demonstrated that the return was great in 2017 but was dismal for 2018 with a 63% reduction. Steve subsequently requested an analysis of the yearly return for all stocks in 2017 and 2018 with the hope of using the information to help his parents find other potential investment opportunities. The results indicated that Sunrun, Inc. (RUN) may have some potential for investment.
### Year 2017 Results
The results of the analysis of the yearly return for 2017 indicated that market performed well overall, with the top three performers as listed below:
- DQ with a total daily volume of 35,796,800 and a return of 199%
- SEDG with a total daily volume of 206,885,200 and a return of 185%
- ENPH with a total daily volume of 221,772,100 and a return of 130%

The three stocks with the lowest returns for 2017 were:
- TERP with a total daily volume of 139,402,800 and a return of -7%
- RUN with a total daily volume of 267,681,300 and a return of 6%
- AY with a total daily volume of 136,070,90 and a return of 9%
### Year 2018 Results
The results of the analysis of the yearly return for 2018 revealed an overall decrease. This could have resulted in part from the impact of "President Donald Trumps' trade war with China, the slowdown in global economic growth and concern that the Federal Reserve was raising interest rates too quickly", as surmised by Gretchen Frazee in her PBS News Hour article titled "_6 factors that fueled the stock market dive in 2018_" (https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018).   

The top three stocks for 2018 were:
- RUN with a total daily volume of 502,757,100 and a return of 84% 
- ENPH with a total daily volume of 607,473,500 and a return of 82%
- VSLR with a total daily volume of 136,539,100 and a return of -4%

The three stocks with the lowest returns for 2018 were:
- DQ with a total daily volume of 107,873,900 and a return of -63%
- JKS with a total daily volume of 158,309,000 and a return of -61%
- SPWR with a total daily volume of 538,024,300 and a return of -45%
## Issues Encountered
Steve was pleased with the analysis and the report generator but relayed that he thought the response time for the output was a little slow. More specifically, he said that the output should return almost instantaneously when the "All Stocks Analysis Button" is clicked. The challenge then, was to streamline the project's visual basic code to make the report run faster. Ultimately, the code was refactored to eliminate separate sub routines, add additional output arrays and modify the looping conventions to to free up some of the memory. A comparison of run times between the original and refactored code revealed that the revisions were a success as illustrated in the screen shots in Appendix 1. Examples of the original and refactored code are also provided for reference, in Appendix 2.
## Conclusion
Some of the benefits of refactoring code are that it can improve efficiency by using less memory and improving logic.decreasing runtimes and eliminating errors in subscripts for calculations as was noted in this project. 
For example: 
- Combining multiple subscripts and enhancing code to use nexted loops to run through the flow once, which decreased the runtime.
- Adding a variable for the ticker index and creating three output arrays for ticker volumes, ticker starting prices and ticker ending prices to improve the logic. 

Disadvantages of refactoring code include that making changes to consolidate subscripts may actually create more problems if the coder isn't careful. For example, the subsript to clear the worksheet could not easily be incorporated into the refactored code without increasing the run time (by this novice). In conclusion, refactoring code should be performed, but it must be done with the utmost care and precision in order for it to be successful. 
## Appendix 1
#### Original Code 2017
![All_Stocks_2017](https://github.com/LleeMcD/Election_Analysis/blob/main/Resources/All_Stocks_2017.png)
#### Refactored Code 2017
![VBA_Challenge_2017](https://github.com/LleeMcD/Election_Analysis/blob/main/Resources/VBA_Challenge_2017.png)
#### Original Code 2018
![All_Stocks_2018](https://github.com/LleeMcD/Election_Analysis/blob/main/Resources/All_Stocks_2018.png)
#### Refactored Code 2018
![VBA_Challenge_2018](https://github.com/LleeMcD/Election_Analysis/blob/main/Resources/VBA_Challenge_2018.png)
## Appendix 2
#### Original Project Code 
Sub yearValueAnalysis()

'Get the year input for the stock analysis from the user

yearValue = InputBox("What year would you like ot run the analyis on?")

End Sub

Sub AllStocksAnalysis()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like ot run the analyis on?")

startTime = Timer

' 1)Format the output sheet on the "All Stocks Analysis" worksheet.

Worksheets("All Stocks Analysis").Activate

Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Initialize an array of all tickers.
Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'3a)Initialize variables for the starting price and ending price.

Dim startingPrice As Single
Dim endingPrice As Single

'3b)Activate the data worksheet

Worksheets(yearValue).Activate

'3c) Find the number of rows to loop over.

  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
'4) Loop through the tickers (outer)

For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0
    
    '5) Loop through rows in the data
    
   Worksheets(yearValue).Activate
   
    For j = 2 To RowCount
    
    '5a) Find total volume for the current ticker.
    
    If Cells(j, 1).Value = ticker Then
    
            totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
    
    '5b) Find the starting price for the current ticker.
    
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
            startingPrice = Cells(j, 6).Value
            
        End If
        
    
    '5c) Find the ending price for the current ticker.
    
     If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
     
        endingPrice = Cells(j, 6).Value
     
     End If
     
    Next j
            
'6) Output the data fo the current ticker.

Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i

endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " + yearValue + ""

End Sub

Sub formatAllStocksAnalysisTable()

'Formatting

Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True

Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

'Formats NumberFormat string for separating digits with commas

Range("B4:B15").NumberFormat = "#,##0"

'Formats a single digit percentage for the return (one digit precision)
                          
Range("C4:C15").NumberFormat = "0.0%"

Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

    If Cells(i, 3) > 0 Then

    'Color the cells green

    Cells(i, 3).Interior.Color = vbGreen

        'Color the cells red

        ElseIf Cells(4, 3) < 0 Then

        Cells(i, 3).Interior.Color = vbRed

    Else

    'Clear the cell color

    Cells(i, 3).Interior.Color = xlNone

    End If
    
    Next i
    
        
End Sub

#### Refactored Project Code
 Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    'Generate a pop-up window to collect a vaule from the end user.
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Add a starting point for the timer.
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
     
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header in row 3
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all stock tickers
    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
       
    
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index.
    Dim tickerIndex As Long
        
    '1b) Create three output arrays.
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    Worksheets(yearValue).Activate
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
        
        '2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
        
            '3a) Increase volume for current ticker.
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
               
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                    
                End If
                
                '3c)Check if the current row is the last row with the selected ticker.
                If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                
                '3d Increase the tickerIndex
                tickerIndex = tickerIndex + 1
                
                End If
                
            Next j
            
                           
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
         For k = 0 To 11
        
         Worksheets("All Stocks Analysis").Activate
             
         Cells(4 + k, 1).Value = tickers(k)
         Cells(4 + k, 2).Value = tickerVolumes(k)
         Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
    
        Next k
    
            'Formatting
             Worksheets("All Stocks Analysis").Activate
             Range("A3:C3").Font.FontStyle = "Bold"
             Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
             Range("B4:B15").NumberFormat = "#,##0"
             Range("C4:C15").NumberFormat = "0.0%"
             Columns("B").AutoFit

             dataRowStart = 4
             dataRowEnd = 15

             For l = dataRowStart To dataRowEnd
        
             If Cells(l, 3) > 0 Then
            
             Cells(l, 3).Interior.Color = vbGreen
            
            Else
        
            Cells(l, 3).Interior.Color = vbRed
            
            End If
        
           Next l
    'Shut the timer off
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

