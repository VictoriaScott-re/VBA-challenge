Attribute VB_Name = "Module1"
 Sub CalculateStockData()
 
    Dim OutputStartRow As Integer
    Dim IsFirstRow As Boolean
    Dim i As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim StockTicker As String
    Dim AnnualChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As LongLong
    Dim GreatestTableIndex As Integer

    For Each ws In Worksheets
        'Assign column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Set initial values
        OutputStartRow = 2
        IsFirstRow = True
        TotalStockVolume = 0

        'Iterate through rows
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            'Check if it's the first row for a ticker
            If IsFirstRow = True Then
                OpeningPrice = ws.Cells(i, 3)
                IsFirstRow = False
            End If

            'Check if next row belongs to a different ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                StockTicker = ws.Cells(i, 1).Value
                ws.Cells(OutputStartRow, 9).Value = StockTicker
                ClosingPrice = ws.Cells(i, 6).Value
                AnnualChange = (ClosingPrice - OpeningPrice)

                'Avoid division by zero
                If OpeningPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = AnnualChange / OpeningPrice
                End If

                'Apply conditional formatting
                If AnnualChange > 0 Then
                    ws.Cells(OutputStartRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(OutputStartRow, 10).Interior.ColorIndex = 3
                End If

                ws.Cells(OutputStartRow, 10).Value = AnnualChange
                ws.Cells(OutputStartRow, 11).Value = PercentChange
                TotalStockVolume = ws.Cells(i, 7).Value + TotalStockVolume
                ws.Cells(OutputStartRow, 12).Value = TotalStockVolume

                IsFirstRow = True
                OutputStartRow = OutputStartRow + 1
                TotalStockVolume = 0
            Else
                TotalStockVolume = ws.Cells(i, 7).Value + TotalStockVolume
            End If
        Next i

        'Format columns
        ws.Columns(9).AutoFit
        ws.Columns(10).AutoFit
        ws.Columns(11).NumberFormat = "0.00%"
        ws.Columns(11).AutoFit
        ws.Columns(12).AutoFit
        ws.Columns(15).AutoFit

        'Find indices for greatest values
        GreatestTableIndex = ws.Cells(Rows.Count, 9).End(xlUp).Row

        For i = 2 To GreatestTableIndex
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(GreatestTableIndex, 11))) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(GreatestTableIndex, 11))) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(GreatestTableIndex, 12))) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 11).Value
            End If
        Next i

        ws.Columns(16).AutoFit
        ws.Range(ws.Cells(1, 17), ws.Cells(2, 17)).NumberFormat = "0.00%"
        ws.Columns(17).AutoFit
        
    Next ws
    
End Sub
