Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double ' Changed to Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double ' Changed to Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim startRow As Long
    Dim endRow As Long

    Dim summarySheet As Worksheet
    Set summarySheet = ThisWorkbook.Sheets.Add
    summarySheet.Name = "Summary"
    summarySheet.Range("A1:C1").Value = Array("Metric", "Ticker", "Value")
    
    maxIncrease = -1
    maxDecrease = 1
    maxVolume = 0

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            startRow = 2
            Dim outputRow As Long
            outputRow = 2

            ws.Range("H:O").ClearContents

            ws.Cells(1, 8).Value = "Ticker"
            ws.Cells(1, 9).Value = "Quarterly Change"
            ws.Cells(1, 10).Value = "Percent Change"
            ws.Cells(1, 11).Value = "Total Stock Volume"

            Do While startRow <= lastRow
                ticker = ws.Cells(startRow, 1).Value
                Dim currentQuarter As String
                currentQuarter = Format(ws.Cells(startRow, 2).Value, "yyyy") & "Q" & Application.WorksheetFunction.RoundUp(Month(ws.Cells(startRow, 2).Value) / 3, 0)
                openPrice = ws.Cells(startRow, 3).Value
                totalVolume = 0

                Do While ws.Cells(startRow, 1).Value = ticker And currentQuarter = Format(ws.Cells(startRow, 2).Value, "yyyy") & "Q" & Application.WorksheetFunction.RoundUp(Month(ws.Cells(startRow, 2).Value) / 3, 0)
                    totalVolume = totalVolume + ws.Cells(startRow, 7).Value
                    closePrice = ws.Cells(startRow, 6).Value
                    startRow = startRow + 1
                    If startRow > lastRow Then Exit Do
                Loop

                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Output the results
                ws.Cells(outputRow, 8).Value = ticker
                ws.Cells(outputRow, 9).Value = quarterlyChange
                ws.Cells(outputRow, 10).Value = percentChange / 100 ' Divide by 100 for percentage formatting
                ws.Cells(outputRow, 10).NumberFormat = "0.00%" ' Format as percentage
                ws.Cells(outputRow, 11).Value = totalVolume

                ' Apply conditional formatting for Quarterly Change
                If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 9).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(outputRow, 9).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                outputRow = outputRow + 1

                ' Find the maximums
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If

                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If

                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
            Loop

            ' Output the maximums on the same worksheet
            ws.Cells(1, 14).Value = "Greatest % Increase"
            ws.Cells(2, 14).Value = maxIncreaseTicker
            ws.Cells(3, 14).Value = maxIncrease / 100 ' Divide by 100 for percentage formatting
            ws.Cells(3, 14).NumberFormat = "0.00%" ' Format as percentage

            ws.Cells(1, 15).Value = "Greatest % Decrease"
            ws.Cells(2, 15).Value = maxDecreaseTicker
            ws.Cells(3, 15).Value = maxDecrease / 100 ' Divide by 100 for percentage formatting
            ws.Cells(3, 15).NumberFormat = "0.00%" ' Format as percentage

            ws.Cells(1, 16).Value = "Greatest Total Volume"
            ws.Cells(2, 16).Value = maxVolumeTicker
            ws.Cells(3, 16).Value = maxVolume
        End If
    Next ws

    ' Output the maximums to the summary sheet
    summarySheet.Cells(2, 1).Value = "Greatest % Increase"
    summarySheet.Cells(2, 2).Value = maxIncreaseTicker
    summarySheet.Cells(2, 3).Value = maxIncrease / 100 ' Divide by 100 for percentage formatting
    summarySheet.Cells(2, 3).NumberFormat = "0.00%" ' Format as percentage

    summarySheet.Cells(3, 1).Value = "Greatest % Decrease"
    summarySheet.Cells(3, 2).Value = maxDecreaseTicker
    summarySheet.Cells(3, 3).Value = maxDecrease / 100 ' Divide by 100 for percentage formatting
    summarySheet.Cells(3, 3).NumberFormat = "0.00%" ' Format as percentage

    summarySheet.Cells(4, 1).Value = "Greatest Total Volume"
    summarySheet.Cells(4, 2).Value = maxVolumeTicker
    summarySheet.Cells(4, 3).Value = maxVolume

    MsgBox "Quarterly Stock Analysis Complete!"
End Sub