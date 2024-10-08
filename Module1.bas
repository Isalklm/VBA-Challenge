Attribute VB_Name = "Module1"
Sub StockAnalysisWithMultipleTickers()

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' Declare variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim TotalVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim i As Long
    Dim QuarterlySheets As Variant
    Dim FirstRow As Long
    Dim LastRowInQuarter As Long
    Dim OutputRow As Long
    Dim FirstTickerRow As Long
    
    ' Variables to store the greatest values
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Define the sheets for each quarter
    QuarterlySheets = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Loop through the sheets to process each quarter
    For Each ws In ThisWorkbook.Sheets(QuarterlySheets)
    
        ' Reset the greatest values for each sheet
        GreatestIncrease = -999999
        GreatestDecrease = 999999
        GreatestVolume = 0

        ' Find the first row (starting data row) and the last row of the sheet
        FirstRow = 2   ' starts in row 2
        LastRowInQuarter = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the first row of the current ticker
        FirstTickerRow = FirstRow
        
        ' Output row for the tickers on the current sheet
        OutputRow = 2

        ' Adding headers
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Volume"
        
        ' Loop through each row of data
        For i = FirstRow To LastRowInQuarter
        
            ' Get the Ticker symbol from Column A
            Ticker = ws.Cells(i, 1).Value
            
            ' Check if the next row contains a different ticker or if it is the last row
            If ws.Cells(i + 1, 1).Value <> Ticker Or i = LastRowInQuarter Then
            
                ' Calculate the quarterly change
                OpeningPrice = ws.Cells(FirstTickerRow, 3).Value ' Opening price from the first row of the ticker
                ClosingPrice = ws.Cells(i, 6).Value ' Closing price from the last row of the ticker
                QuarterlyChange = ClosingPrice - OpeningPrice
                
                ' Calculate the percent change
                If OpeningPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
                
                ' Calculate total volume for the ticker
                TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(FirstTickerRow, 7), ws.Cells(i, 7)))
                
                ' Output the results for the current sheet
                ws.Cells(OutputRow, 10).Value = Ticker               ' Column J: Ticker symbol
                ws.Cells(OutputRow, 11).Value = QuarterlyChange      ' Column K: Quarterly Change
                ws.Cells(OutputRow, 12).Value = PercentChange        ' Column L: Percent Change
                ws.Cells(OutputRow, 13).Value = TotalVolume          ' Column M: Total Stock Volume
                
                ' Find the greatest percent increase
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                
                ' Find the greatest percent decrease
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                
                ' Find the greatest total volume
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Set the new first row for the next ticker
                FirstTickerRow = i + 1
            End If
        Next i

        ' Output the greatest values into specific cells (with headers) on each sheet
        With ws
            ' Add headers for columns O, P, Q for each sheet's summary box
            .Cells(1, 15).Value = "Description"   ' Column O header
            .Cells(1, 16).Value = "Ticker"        ' Column P header
            .Cells(1, 17).Value = "Value"         ' Column Q header

            ' Output the greatest values with correct formatting
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = GreatestIncreaseTicker
            .Cells(2, 17).Value = GreatestIncrease   ' Store the percentage value
            .Cells(2, 17).NumberFormat = "0.00%"     ' Format as percentage

            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = GreatestDecreaseTicker
            .Cells(3, 17).Value = GreatestDecrease  ' Store the percentage value
            .Cells(3, 17).NumberFormat = "0.00%" ' Format as percentage

            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = GreatestVolumeTicker
            .Cells(4, 17).Value = GreatestVolume ' Store the value in scientific notation
            .Cells(4, 17).NumberFormat = "0.00E+00" ' Format as scientific notation
        End With

        ' Apply conditional formatting for Quarterly Change (Column K)
        For Each cell In ws.Range("K2:K" & OutputRow - 1)
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4       ' Green color for positive values
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3       ' Red color for negative values
            Else
                cell.Interior.ColorIndex = 0       ' No color for zero values
            End If
        Next cell

        ' Apply conditional formatting for Percent Change (Column L)
        For Each cell In ws.Range("L2:L" & OutputRow - 1)
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4       ' Green color for positive values
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3       ' Red color for negative values
            Else
                cell.Interior.ColorIndex = 0       ' No color for zero values
            End If
        Next cell
          
    Next ws
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True

End Sub

