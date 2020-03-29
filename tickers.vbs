Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call tickerValue
    Next
    Application.ScreenUpdating = True
End Sub
Sub tickerValue()
' setting value for the ticker name
Dim ticker As String
' Where we find the last row of our data
Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
' Where we want to output my data
Dim TickerSumRow As Double
TickerSumRow = 2
' setting the initial open price for the pct chng
Dim initopenPrice As Double
initopenPrice = Cells(2, 3)
' setting up structure for close price
Dim closePrice As Double
' Setting up structure for yearly percent change
Dim yearlyPcntChng As Double
'this sets up the variable for the ticker yearly price change
Dim yearlyPriceChng As Double
'initializing total
Dim TotalVolume As Double
TotalVolume = Cells(2, 7)
Dim GreatestPcntInc As Double
Cells(1, 9) = ("Ticker")
Cells(1, 10) = ("Percent Change")
Cells(1, 11) = ("Yearly Share Change")
Cells(1, 12) = ("Stock Total Volume")
Cells(2, 14) = ("Greatest Percent Increase")
Cells(2, 15) = GreatestPcntInc
' when the ticker changes capture the ticker and info
    For i = 2 To lastrow
       'Calculating total stock volume
        TotalVolume = Cells(i, 7) + TotalVolume
        
        'setting up the error function
        On Error Resume Next
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'Grabbing ticker Value
                Cells(TickerSumRow, 9).Value = Cells(i, 1).Value
                'Grabbing the stocks closing price
                closePrice = Cells(i, 6).Value
                'Calculating the tickers percent change
                yearlyPcntChng = ((closePrice - initopenPrice) / initopenPrice)
                'Assigning the Yearly Percent to a position
                Cells(TickerSumRow, 10).Value = yearlyPcntChng
                'Calculating yearlyPriceChng
                yearlyPriceChng = (closePrice - initopenPrice)
                'Assigining location for yearlyPriceChange Values
                Cells(TickerSumRow, 11) = yearlyPriceChng
                'Assigning location for total stock volume
                Cells(TickerSumRow, 12) = TotalVolume
                'grabbing the stocks initial price
                initopenPrice = Cells(i + 1, 3).Value
                'once last row is done need to reset for next instance
                GreatestPcntInc = WorksheetFunction.Max(yearlyPcntChng)
                TotalVolume = Cells(i + 1, 7)
                lastrow = 0
                TickerSumRow = TickerSumRow + 1
    
             End If
             
                If Cells(i, 11).Value > 0 Then
                    Cells(i, 11).Interior.ColorIndex = 4
                Else: Cells(i, 11).Interior.ColorIndex = 3
                End If
                    
                
                    
                    
        
    Next i
        
               
End Sub


