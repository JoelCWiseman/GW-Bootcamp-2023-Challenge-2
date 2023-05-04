Attribute VB_Name = "Module2"
Sub stock_data()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim SummaryTableRow As Long
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentIncrease As String
    Dim tickerMaxPercentDecrease As String
    Dim tickerMaxTotalVolume As String
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Values"
        
        ' Find last row of data
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
               
        SummaryTableRow = 2
        
        ' Loop through all rows of data
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ' Set Closing Price
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Set Yearly Change
                OpenPrice = ws.Cells(i - WorksheetFunction.CountIf(ws.Range("A2:A" & i - 1), Ticker), 3).Value
                YearlyChange = ClosePrice - OpenPrice
                
                ' Set Percent Change
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                
                ' Set Total Volume
                TotalVolume = WorksheetFunction.Sum(ws.Range("G" & (i - WorksheetFunction.CountIf(ws.Range("A2:A" & i - 1), Ticker)) & ":G" & i))
                
                ' Output results to Summary Table
                ws.Range("J" & SummaryTableRow).Value = Ticker
                ws.Range("K" & SummaryTableRow).Value = YearlyChange
                ws.Range("L" & SummaryTableRow).Value = PercentChange
                ws.Range("L" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("M" & SummaryTableRow).Value = TotalVolume
                
                ' Format Yearly Change cell with conditional formatting
                If YearlyChange > 0 Then
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                
                ' Increment Summary Table row
                SummaryTableRow = SummaryTableRow + 1
                
                ' Reset variables
                YearlyChange = 0
                PercentChange = 0
                TotalVolume = 0
                
            Else
                ' Add to Total Volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Add conditional formatting to Column K
        ws.Range("K2:K" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ws.Range("K2:K" & LastRow).FormatConditions(ws.Range("K2:K" & LastRow).FormatConditions.Count).SetFirstPriority
        ws.Range("K2:K" & LastRow).FormatConditions(1).Interior.Color = 4
        
        ws.Range("K2:K" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range("K2:K" & LastRow).FormatConditions(ws.Range("K2:K" & LastRow).FormatConditions.Count).SetFirstPriority
        ws.Range("K2:K" & LastRow).FormatConditions(1).Interior.Color = 3
        
        ' Find the greatest % increase, greatest % decrease and greatest total volume
        For i = 2 To LastRow
            ' Check for greatest % increase
            If ws.Cells(i, 12).Value > maxPercentIncrease Then
                maxPercentIncrease = ws.Cells(i, 12).Value
                tickerMaxPercentIncrease = ws.Cells(i, 10).Value
            End If

            ' Check for greatest % decrease
            If ws.Cells(i, 12).Value < maxPercentDecrease Then
                maxPercentDecrease = ws.Cells(i, 12).Value
                tickerMaxPercentDecrease = ws.Cells(i, 10).Value
            End If

            ' Check for greatest total volume
            If ws.Cells(i, 13).Value > maxTotalVolume Then
                maxTotalVolume = ws.Cells(i, 13).Value
                tickerMaxTotalVolume = ws.Cells(i, 10).Value
            End If
        Next i
        
        ' Output the results to the summary table
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q2").Value = tickerMaxPercentIncrease
        ws.Range("Q3").Value = tickerMaxPercentDecrease
        ws.Range("Q4").Value = tickerMaxTotalVolume
        ws.Range("R2").Value = Format(maxPercentIncrease, "0.00%")
        ws.Range("R3").Value = Format(maxPercentDecrease, "0.00%")
        ws.Range("R4").Value = maxTotalVolume
    Next ws
End Sub
