Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim lastRow As Long
    Dim outputRow As Long
    Dim i As Long

    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Add headers if they do not exist
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Quarterly Change ($)"
        ws.Cells(1, 12).Value = "Percent Change (%)"
        
        ws.Cells(1, 14).Value = "Metric"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        outputRow = 2
        maxIncrease = -1
        maxDecrease = 1
        maxVolume = 0
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value ' Assuming Volume is in Column G
            openPrice = ws.Cells(i, 3).Value ' Assuming Open Price is in Column C
            closePrice = ws.Cells(i, 6).Value ' Assuming Close Price is in Column F
            
            ' Calculate metrics
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If

            ' Output results
            ws.Cells(outputRow, 9).Value = ticker ' Assuming output starts in Column I
            ws.Cells(outputRow, 10).Value = volume
            ws.Cells(outputRow, 11).Value = quarterlyChange
            ws.Cells(outputRow, 12).Value = percentChange

            ' Apply conditional formatting
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, 11).Interior.Color = vbGreen
            ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, 11).Interior.Color = vbRed
            End If

            If percentChange > 0 Then
                ws.Cells(outputRow, 12).Interior.Color = vbGreen
            ElseIf percentChange < 0 Then
                ws.Cells(outputRow, 12).Interior.Color = vbRed
            End If

            ' Track max values
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = ticker
            End If

            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = ticker
            End If

            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeTicker = ticker
            End If

            outputRow = outputRow + 1
        Next i

        ' Output the max values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = maxIncrease

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = maxDecrease

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(4, 16).Value = maxVolume

        ' Apply formatting to headers
        ws.Cells(1, 9).Font.Bold = True
        ws.Cells(1, 10).Font.Bold = True
        ws.Cells(1, 11).Font.Bold = True
        ws.Cells(1, 12).Font.Bold = True
        ws.Cells(1, 14).Font.Bold = True
        ws.Cells(1, 15).Font.Bold = True
        ws.Cells(1, 16).Font.Bold = True
    Next ws
End Sub

