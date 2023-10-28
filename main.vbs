Option Explicit
Sub main()
 ' year start date
 Const startDate = "0102"
 ' year end date
 Const endDate = "1231"
 ' year start price per ticker
 Dim yearStartPrice As Double
 ' year end price per ticker
 Dim yearEndPrice As Double
 ' year end price change per ticker
 Dim yearPriceChange As Double
 ' year end percentage change per ticker
 Dim yearPercentageChange As Double
 ' year end total volume per ticker
 Dim totalVolume As Double
 ' current worksheet in process
 Dim ws As Worksheet
 
  'Loop through each worksheet
 For Each ws In Worksheets
 
    Dim geatestPctIncrease, geatestPctDecrease As Double
    Dim geatestTotalVolume As LongLong
    Dim ticker_geatestPctIncrease, ticker_geatestPctDecrease, ticker_geatestTotalVolume As String
    ' set initial value before looping each worksheet
    geatestPctIncrease = 0
    geatestPctDecrease = 0
    geatestTotalVolume = 0
  
    'each ticker yearly summary hearder output
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
  
    'last row of each worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'row of last ticket yearly summary output
    Dim counter As Integer
    'start from 2nd row
    counter = 2
    'Loop Through each row
    Dim r As Long
    
    For r = 2 To lastRow
        ' metrics for each row
        Dim ticker As String
        Dim marketDate As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim volume As LongLong
                                        
        ticker = ws.Cells(r, 1).Value
        marketDate = ws.Cells(r, 2).Value
        openPrice = ws.Cells(r, 3).Value
        closePrice = ws.Cells(r, 6).Value
        volume = ws.Cells(r, 7)
        
        ' compare the last 4 characters of market date with static constant variables: startDate and endDate to find start date and end date
        If Right$(marketDate, 4) = startDate Then
            yearStartPrice = openPrice
        ElseIf Right$(marketDate, 4) = endDate Then
            yearEndPrice = closePrice
        End If
        
        ' total volume add up by each date
        totalVolume = totalVolume + volume
        
        ' if both yearly start and end price found
        ' then output the ticker summary and then reset the metrics and move on to next ticker
        If yearStartPrice <> 0 And yearEndPrice <> 0 Then
            
            'Yearly Change by the ticker
            yearPriceChange = yearEndPrice - yearStartPrice
            
            'Yearly Percentage change by the ticker
            yearPercentageChange = Round(yearPriceChange / yearStartPrice, 4)
            
            'output the year end summary info for each ticker
            ws.Cells(counter, 9) = ticker
            ws.Cells(counter, 10) = yearPriceChange

            'conditional format for price change; red for negative and green for positive
            If yearPriceChange > 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 4 'green
            ElseIf yearPriceChange < 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 3 'red
            End If

            'yearly percentage change output by the ticker
            ws.Cells(counter, 11) = yearPercentageChange
            ws.Cells(counter, 11).NumberFormat = "0.00%"
            
            'yearly total volume output by the ticker
            ws.Cells(counter, 12) = totalVolume
            ws.Cells(counter, 12).NumberFormat = "#,##0"
                    
            
            'increment the counter for next ticker
            counter = counter + 1
            
            'record greatest metrices among the current worksheet
            If yearPercentageChange > geatestPctIncrease Then
                geatestPctIncrease = yearPercentageChange
                ticker_geatestPctIncrease = ticker
            End If
            
            If yearPercentageChange < geatestPctDecrease Then
                geatestPctDecrease = yearPercentageChange
                ticker_geatestPctDecrease = ticker
            End If
            
            If totalVolume > geatestTotalVolume Then
                geatestTotalVolume = totalVolume
                ticker_geatestTotalVolume = ticker
            End If
                        
            'reset the metrics for next ticker
            yearStartPrice = 0
            yearEndPrice = 0
            yearPriceChange = 0
            yearPercentageChange = 0
            totalVolume = 0
        End If
     Next r
     
      'Greatest info output
     ws.Cells(1, 16) = "Ticker"
     ws.Cells(1, 17) = "Value"
        
     ws.Cells(2, 15) = "Greatest % Increase"
     ws.Cells(2, 16) = ticker_geatestPctIncrease
     ws.Cells(2, 17) = geatestPctIncrease
     ws.Cells(2, 17).NumberFormat = "0.00%"
     
     ws.Cells(3, 15) = "Greatest % Decrease"
     ws.Cells(3, 16) = ticker_geatestPctDecrease
     ws.Cells(3, 17) = geatestPctDecrease
     ws.Cells(3, 17).NumberFormat = "0.00%"
     
     ws.Cells(4, 15) = "Greatest Total Volume"
     ws.Cells(4, 16) = ticker_geatestTotalVolume
     ws.Cells(4, 17) = geatestTotalVolume
     ws.Cells(4, 17).NumberFormat = "#,##0"
     
     ' auto fit for output range
     ws.Range("I1:L1").Columns.AutoFit
     ws.Range("O1:Q4").Columns.AutoFit
     
Next
End Sub
