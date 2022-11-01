# VBA code refactoring to increased performance of the code for Steve.

## Overview of Project
We have produced a VBA Workbook for [**Steve**](https://en.wikipedia.org/wiki/Steve_Ballmer) to quickly analyze a list of 12 green energy stocks and produce a dashboard to display performance and traded volumes for years 2017 and 2018. Steve loved our workbook and has requested to expand this analysis to the entire stock market and additional years. We will assume [**S&P 500 Index**](https://en.wikipedia.org/wiki/S%26P_500) which contains around 500 stock names and will undoubtedly slow down our code performance due to the volume of data to loop through and larger output. 

## Purpose
The purpose of this project is to refactor and streamline our VBA code for stock analysis to make it more efficient and flexible which will allow us to analyze larger datasets.

## Analysis and Challenges
Given that our dataset will increase almost 50x the main challenge for us is to refactor the code to loop through the data just once instead of using multiple loops.

### Analysis and Code Comparison:
<details>
 <summary>Click here to see Original Code</summary>

  ### Original Code [VBA_Challenge_original.vbs](https://github.com/vkbt/stock-analysis/blob/main/VBA_Challenge_original.vbs)
  ```vba
 
 Sub AllStockAnalysis()

'define startTime and endTime as Singles

Dim startTime As Single
Dim endTime  As Single

'define variable yearvalue and create input box with question to user which year's information will be used in this analysis

yearvalue = InputBox("What year would you like to run the analysis on?")

'start timer

startTime = Timer

'select "All Stocks Analysis Worksheet"

Worksheets("All Stocks Analysis").Activate

'name and format headers in the worksheet

Range("A1").Value = "All Stocks (" + yearvalue + ")"
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit


Cells(3, 1).Value = "Ticker"

Cells(3, 2).Value = "Total Daily Volume"

Cells(3, 3).Value = "Return"

'initialize array of 12 tickers and assign String type to the tickers

Dim tickers(11) As String

'define ticker names

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
    
'define startingPrice and endingPrice as Double

Dim startingPrice As Double
Dim endingPrice As Double

'select worksheet with data

Worksheets(yearvalue).Activate

'define rowStart and RowEnd

rowStart = 2
RowEnd = Cells(Rows.Count, "A").End(xlUp).Row

'Start loop for tickers

For i = 0 To 11

ticker = tickers(i)

'define totalVolume and set it at 0

totalVolume = 0

'select worksheet with data

Worksheets(yearvalue).Activate

'start loop for totalvolume by 1) defining "j" using rowStart and rowEnd

For j = rowStart To RowEnd

    'create IF rule for totalVolume that changes if ticker name is changing
    
    If Cells(j, 1).Value = ticker Then
        
    totalVolume = totalVolume + Cells(j, 8).Value
       
    End If
    
    'create IF rule for startingPrice
    
    If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
    
    startingPrice = Cells(j, 6)
        
    End If
    
    'create IF rule for endingPrice
    
    If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
    
    endingPrice = Cells(j, 6)
    
    End If
    
'end j loop

Next j

'select "All Stocks Analysis Worksheet"

Worksheets("All Stocks Analysis").Activate

'populate list of tickers

Cells(4 + i, 1).Value = ticker

'populate totalvolume result

Cells(4 + i, 2).Value = totalVolume

'populate value result

Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

'formatting

If Cells(4 + i, 3) > 0 Then

Cells(4 + i, 3).Interior.Color = vbGreen

ElseIf Cells(4 + i, 3) < 0 Then

    Cells(4 + i, 3).Interior.Color = vbRed

Else

    Cells(4 + i, 3).Interior.Color = xlNone

End If       

Next i
endTime = Timer
MsgBox "Original code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub
  ```
</details>

<details>
 <summary>Click here to see Refactored Code</summary>
 
  ### Refactored code [VBA_Challenge_refactored.vbs](https://github.com/vkbt/stock-analysis/blob/main/VBA_Challenge_refactored.vbs) 
```vba
 
 Sub AllStocksAnalysisRefactored()
    Application.ScreenUpdating = False
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(11) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrice(11) As Single
    Dim tickerEndingPrice(11) As Single
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            If Cells(i, 1) = tickers(tickerIndex) Then tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
          
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then tickerStartingPrice(tickerIndex) = Cells(i, 6)
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then tickerEndingPrice(tickerIndex) = Cells(i, 6)
        '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
            
        'End If
     
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    Cells(4 + i, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickerVolumes(i)
    Cells(i + 4, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "Refactored code + Screen Update = OFF ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub
```
 
</details>


#### Code Performance and Findings:
**Final Workbook:** [VBA_Challenge.xlsm](https://github.com/vkbt/stock-analysis/blob/main/VBA_Challenge.xlsm)
 
 To measure performance of our original code we have implemented a timer routine within our code and took notes of code run times running 2017 data as well as 2018.

#### 2017 vs 2018 - [Original code](https://github.com/vkbt/stock-analysis/blob/main/VBA_Challenge_original.vbs) with multiple loops
<p float="left">
 <img src="https://github.com/vkbt/stock-analysis/blob/main/resources/Original%20code%202017.png" width=40% height=40%>
 <img src="https://github.com/vkbt/stock-analysis/blob/main/resources/Original%20code%202018.png" width=40% height=40%>
 </p>


By refactoring our code and streamlining **For** loop statements into one loop we were able to drastically improve our code performance which is evident if we compare run times of our original code vs. our newly refactored code.

#### 2017 vs 2018 - [Refactored code](https://github.com/vkbt/stock-analysis/blob/main/VBA_Challenge_refactored.vbs) with one loop
 <p float="left">
<img src="https://github.com/vkbt/stock-analysis/blob/main/resources/Refactored%20code%202017.png" width=40% height=40%>
<img src="https://github.com/vkbt/stock-analysis/blob/main/resources/Refactored%20code%202018.png" width=40% height=40%>
</p>

## Conclusion and Challenges
While the original code is well written and allows us to analyze smaller subsets of data the reality is that while we are very happy with our original code it will not be as robust in its performance analyzing larger amounts of stock market data.
We have refactored the code to restructure multiple For loops into one more complex For loop allowing us to increase codes performance, decrease run time and optimize its structure.
As for disadvantages refactoring this code produced several new problems that took time to debug, find errors and fix them.

## Bonus - Turn off Screen Updates
While exploring different VBA resources online we were able to find a VBA switch to turn off screen updates while the code is running.
 
**Application.ScreenUpdating = False**
 
By placing this line in the beginning of our subroutine it will significantly reduce script run time on top of our current refactoring.
 

