Attribute VB_Name = "Module1"
Sub summarizeStocks()

    ' declare variables
    Dim currentTicker As String
    Dim nextTicker As String
    Dim counter As Integer
    Dim totalStockVolume As LongLong
    Dim openingPrice, currentClosingPrice, percentChange As Double
        
    ' iterate over all worksheets to calculate "Yearly Change", "Percent Change", "Total Stock Volume"
    For Each ws In Worksheets

        lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
                
        ' initialize variables
        counter = 2  ' initialize row counter for storing unique tickers
        currentTicker = ws.Range("A2").Value ' first value to evaluate
        ws.Range("I1").Value = "Ticker"  ' Add "Ticker" header
        openingPrice = ws.Range("C2").Value ' initialize first ticker's "opening price"
        currentClosingPrice = 0
        ws.Range("J1").Value = "Yearly Change" ' Add "Yearly Change" header
        percentChange = 0
        ws.Range("K1").Value = "Percent Change" ' Add "Percent Change" header
        totalStockVolume = 0
        ws.Range("L1").Value = "Total Stock Volume"  ' Add "Total Stock Volume" header
        
        For i = 3 To lastRow
            
            ' initialize value
            nextTicker = ws.Range("A" & i).Value
            
            If currentTicker <> nextTicker Then
            
                ' set current unique Ticker value to cell
                ws.Range("I" & counter).Value = currentTicker
                
                ' calculate and set current "yearly change"
                ws.Range("J" & counter).Value = openingPrice - currentClosingPrice
                If ws.Range("J" & counter).Value < 0 Then
                    ' Set the Cell Colors to Red
                    ws.Range("J" & counter).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & counter).Interior.ColorIndex = 4
                End If
                               
                ' calulate and set current "percent change"
                If openingPrice <> 0 Then
                    percentChange = currentClosingPrice / openingPrice
                    
                Else
                    percentChange = 0
                End If
                
                ws.Range("K" & counter).Value = Format(percentChange, "Percent")
                ' set current ticker's "total stock volume"
                ws.Range("L" & counter).Value = totalStockVolume
                
                ' set the variable for next unique value evaluation
                currentTicker = nextTicker
                totalStockVolume = ws.Range("G" & i).Value
                openingPrice = ws.Range("C" & i).Value
                currentClosingPrice = ws.Range("F" & i).Value
                counter = counter + 1
                
            Else
                totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
                currentClosingPrice = ws.Range("F" & i).Value
            End If

        Next i
    
    Next ws


    Dim greatestPercentIncrease As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolume As LongLong
    Dim greatestTotalVolumeTicker As String
    
    ' Calculate summary data
    For Each ws2 In Worksheets

        lastRow = ws2.Range("A" & ws2.Rows.Count).End(xlUp).Row
        
        For j = 2 To lastRow
            
            If ws2.Range("K" & j).Value > greatestPercentIncrease Then
                greatestPercentIncreaseTicker = ws2.Range("A" & j)
                greatestPercentIncrease = ws2.Range("K" & j).Value
            End If
            
            If ws2.Range("K" & j).Value < greatestPercentDecrease Then
                greatestPercentDecreaseTicker = ws2.Range("A" & j)
                greatestPercentDecrease = ws2.Range("K" & j).Value
            End If
            
            If ws2.Range("L" & j).Value > greatestTotalVolume Then
                greatestTotalVolumeTicker = ws2.Range("A" & j)
                greatestTotalVolume = ws2.Range("L" & j).Value
            End If
        
        Next j
            
        
    Next ws2
    
    ' Display summary data
    ' Column Headers labels
    Worksheets("2014").Range("P1").Value = "Ticker"
    Worksheets("2014").Range("Q1").Value = "Value"
    
    ' set rows of summary data
    Worksheets("2014").Range("O2").Value = "Greatest % Increase"
    Worksheets("2014").Range("P2").Value = greatestPercentIncreaseTicker
    Worksheets("2014").Range("Q2").Value = greatestPercentIncrease
    
    Worksheets("2014").Range("O3").Value = "Greatest % Decrease"
    Worksheets("2014").Range("P3").Value = greatestPercentDecreaseTicker
    Worksheets("2014").Range("Q3").Value = greatestPercentDecrease
    
    Worksheets("2014").Range("O4").Value = "Greatest Total Volume"
    Worksheets("2014").Range("P4").Value = greatestTotalVolumeTicker
    Worksheets("2014").Range("Q4").Value = greatestTotalVolume


End Sub
