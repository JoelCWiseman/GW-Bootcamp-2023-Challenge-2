Attribute VB_Name = "Module1"
Sub stock_data()

Dim ws As Worksheet
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Double
Dim Ticker As String
Dim Yearly_Change As Double
Dim percent_change As Double
Dim total_stock_value As Double
Dim summary_table_row As Integer
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double

summary_table_row = 2
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("R4").NumberFormat = "0"

    ws.Range("J1:M1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

    For i = 2 To rowCount

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        Yearly_Change = Yearly_Change + ws.Cells(i, 6).Value - ws.Cells(i, 3).Value

        total_stock_value = total_stock_value + ws.Cells(i, 7).Value

        percentChange = (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value

        ws.Range("j" & summary_table_row).Value = ws.Cells(i, 1).Value

        ws.Range("k" & summary_table_row).Value = Yearly_Change

        ws.Range("l" & summary_table_row).Value = Format(percentChange, "Percent")
        
        ws.Range("m" & summary_table_row).Value = total_stock_value

        summary_table_row = summary_table_row + 1

        Yearly_Change = 0

        total_stock_value = 0

      Else

        Ticker = ws.Cells(i, 1).Value

        Yearly_Change = Yearly_Change + ws.Cells(i, 6).Value - ws.Cells(i, 3).Value

        total_stock_value = total_stock_value + ws.Cells(i, 7).Value

        If ws.Range("K" & i).Value < 0 Then
          ws.Range("K" & i).Interior.ColorIndex = 3 'Red
        ElseIf ws.Range("K" & i).Value > 0 Then
          ws.Range("K" & i).Interior.ColorIndex = 4 'Green
        End If

      End If
      
    Next i

    summary_table_row = 2
    ws.Range("Q2").Value = WorksheetFunction.Index(ws.Range("J:J"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0))
    ws.Range("Q3").Value = WorksheetFunction.Index(ws.Range("J:J"), WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L:L")), ws.Range("L:L"), 0))
    ws.Range("Q4").Value = WorksheetFunction.Index(ws.Range("J:J"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M:M")), ws.Range("M:M"), 0))

    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"

    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"

    greatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("R2").Value = greatestIncrease

    greatestDecrease = WorksheetFunction.Min(ws.Range("L:L"))
    ws.Range("R3").Value = greatestDecrease

    greatestVolume = WorksheetFunction.Max(ws.Range("M:M"))
    ws.Range("R4").Value = greatestVolume

Next ws

End Sub

