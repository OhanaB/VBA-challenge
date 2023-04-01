Attribute VB_Name = "Module1"
Sub STOCK_MARKET_ANALYSIS_OHANA()
    'LOOPING TROUGH EACH WK SHEET
    For Each ws In Worksheets
        'SUMMARY TABLE COLLUMN NAME
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'BONUS TABLE COLLUMN NAME
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'VARIABLE FOR SUMMARY TABLE
        Dim Ticker As String
        Dim Value As Integer
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        'VARIABLES FOR BONUS TABLE
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestVolume As Double
        GreatestVolume = 0
        Dim GreatestTicker As String

        'IDENTIFYING THE LAST ROW OF THE DATA
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'LOOPING TROUGH ALL ROWS
        For I = 2 To LastRow
            'IDENTIFYING IF STILL ON THE SAME TICKER
            If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then
                'Set the ticker
                Ticker = ws.Cells(I, 1).Value
                'Set the opening price
                OpeningPrice = ws.Cells(I, 3).Value
                'Reset the total volume
                TotalVolume = 0
            End If
            
            'ADDING TO THE TOTAL VOLUME
            TotalVolume = TotalVolume + ws.Cells(I, 7).Value
            'ARE WE ON THE SAME TICKER?
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                'SETTING THE CLOSING PRICE
                ClosingPrice = ws.Cells(I, 6).Value
                'CALCULATING YEARLY AND PERCENTAGE CHANGE
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = YearlyChange / OpeningPrice
                Else
                    PercentChange = 0
                End If
                ' PRINT TICKER, YEARLY CHANGE, PERCENT CHANGE AND TOTAL VOLUME ON THE SUMMARY TABLE
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                'PERCENT CHANGE CONDITIONAL FORMARTING
                If YearlyChange > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 0
                End If
               If PercentChange > 0 Then
                'YEARLY CHANGE CONDITIONAL FORMARTING

               ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
           ElseIf PercentChange < 0 Then
           ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
           Else
           ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 0
           End If
       
                'INCREASE THE ROW COUNTER
                SummaryTableRow = SummaryTableRow + 1
                     
                 If YearlyChange > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 0
                End If
                
                'INDENTIFYING THE GREATEST % INCREASE FOR TICKER AND VALUE
                 'INDENTIFYING THE GREATEST % DECREASE FOR TICKER AND VALUE
                'INDENTIFYING THE GREATEST TOTAL VOLUEME FOR TIKER AND VALUE

                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestTickerincrease = Ticker
                
                ElseIf PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestTickerDecrease = Ticker
                    End If
                    
           If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                GreatestVolumeTicker = Ticker
                End If
              End If
        Next I
        
        
        'PRINT THE VALUES FOR THE BONUS TABLE
        ws.Range("P2").Value = GreatestTickerincrease
        ws.Range("Q2").Value = Format(GreatestIncrease, "0.00%")
        ws.Range("P3").Value = GreatestTickerDecrease
        ws.Range("Q3").Value = Format(GreatestDecrease, "0.00%")
        ws.Range("P4").Value = GreatestVolumeTicker
        ws.Range("Q4").Value = Format(GreatestVolume, "#,##0")
   Next ws

End Sub


