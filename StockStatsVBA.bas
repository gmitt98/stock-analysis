Attribute VB_Name = "Module1"
Sub StockStats()

    Dim ws As Worksheet
    
    'Loop that does all the calculations on each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        ' Find the last row containing a stock price
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.count, "A").End(xlUp).Row
        
        'Declare the variables I need to use and initialize the stock var and start price so that each sheet starts clean
        Dim stock As String
        stock = ws.Cells(2, 1).Value
        Dim startPrice As Double
        startPrice = ws.Cells(2, 3)
        Dim endPrice As Double
        Dim yearDelta As Double
        Dim percentDelta As Double
        Dim rowToWrite As Long
        Dim volume As Double
        volume = 0
        
        'Declare the variables I'm using to keep track of the extremes for the 3 row summary table and initialize them to start each worksheet over
        Dim maxPercentStock As String
        maxPercentStock = ""
        Dim maxPercentValue As Double
        maxPercentValue = 0
        Dim minPercentStock As String
        minPercentStok = ""
        Dim minPercentValue As Double
        minPercentValue = 0
        Dim maxVolumeStock As String
        maxVolumeStock = ""
        Dim maxVolumeValue As Double
        maxVolumeValue = 0

        
        'Write the headers for my output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        'Write the row labels for the last table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Set the first row we're going to write the output of stocks into which goes under my headers
        rowToWrite = 2
        
        'Iterate through the rows and add up our stats
        For i = 2 To lastRow + 1
            'Check if the row we are on is the same as the stock we expect it to be, and if it changed, reset volume and capture stuff
            'Because this is the end of this stock's range, we calculate the stats and write it out, changing the cell number format as needed
            If ws.Cells(i, 1).Value <> stock Then
                endPrice = ws.Cells(i - 1, 6).Value
                yearDelta = endPrice - startPrice
                percentDelta = yearDelta / startPrice
                ws.Cells(rowToWrite, 9).Value = stock
                ws.Cells(rowToWrite, 10).Value = yearDelta
                ws.Cells(rowToWrite, 10).NumberFormat = "0.."
                
                'This If statement block changes the cell color based on the value contained, with red for negative and green for postiive
                If yearDelta < 0 Then
                    ws.Cells(rowToWrite, 10).Interior.ColorIndex = 3
                ElseIf yearDelta > 0 Then
                    ws.Cells(rowToWrite, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowToWrite, 10).Interior.ColorIndex = xlNone
                End If
                
                'Back to writing out out results now
                ws.Cells(rowToWrite, 11).Value = percentDelta
                ws.Cells(rowToWrite, 11).NumberFormat = "0.00%"
                ws.Cells(rowToWrite, 12).Value = volume
                ws.Cells(rowToWrite, 12).NumberFormat = "0"
                
                'check against current greatest percent increase, decrease, and greatest volume seen so far
                If percentDelta > maxPercentValue Then
                    maxPercentValue = percentDelta
                    maxPercentStock = stock
                End If
                If percentDelta < minPercentValue Then
                    minPercentValue = percentDelta
                    minPercentStock = stock
                End If
                If volume > maxVolumeValue Then
                    maxVolumeValue = volume
                    maxVolumeStock = stock
                End If
                
                'Now we get the new stock ticker value and new start price and reset volume
                startPrice = ws.Cells(i, 3).Value
                stock = ws.Cells(i, 1).Value
                volume = 0
                rowToWrite = rowToWrite + 1
            End If
            'Stock ticker hasn't changed, so we just add up the volume and move on
            volume = volume + ws.Cells(i, 7).Value
            
        Next i
        
    'Now we write the max values out in the small summary table, and change their number format
    ws.Cells(2, 16).Value = maxPercentStock
    ws.Cells(2, 17).Value = maxPercentValue
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = minPercentStock
    ws.Cells(3, 17).Value = minPercentValue
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = maxVolumeStock
    ws.Cells(4, 17).Value = maxVolumeValue
    ws.Cells(4, 17).NumberFormat = "0"
    
    Next ws
    
End Sub

