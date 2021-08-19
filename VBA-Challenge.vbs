Sub The_VBA_of_Wall_Street()

    For Each ws In Worksheets

        Dim Ticker As String
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim YearlyOpeningPrice As Double
        Dim YearlyClosingPrice As Double
        Dim YearlyPriceChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim YearlyPercentageChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
    
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"


        For i = 2 To lastRow

            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                TotalStockVolume = 0

                YearlyOpeningPrice = ws.Range("C" & PreviousAmount)
                YearlyClosingPrice = ws.Cells(i, 6)
                YearlyPriceChange = YearlyClosingPrice - YearlyOpeningPrice
                ws.Range("J" & SummaryTableRow).Value = YearlyPriceChange

            If YearlyOpeningPrice = 0 Then
                YearlyPercentageChange = 0

            Else
                YearlyOpeningPrice = ws.Range("C" & PreviousAmount)
                YearlyPercentageChange = YearlyPriceChange / YearlyOpeningPrice

            End If

            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            ws.Range("K" & SummaryTableRow).Value = YearlyPercentageChange

            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                
            Else
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            
            End If

            SummaryTableRow = SummaryTableRow + 1
            PreviousAmount = i + 1
            
            End If

        Next i

        For i = 2 To lastRow

            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            
            End If

            If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            
            End If

            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
        
            End If

        Next i

            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Columns("I:Q").AutoFit

    Next ws

End Sub

        
