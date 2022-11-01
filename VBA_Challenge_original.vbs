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
