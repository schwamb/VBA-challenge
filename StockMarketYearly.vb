Dim LastRow As LongLong
Dim UniqueTicker As String
Dim SummaryTable As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVol As LongLong

Sub StockMarket()
SummaryTable = 2
'loops through the worksheets
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Find the Final Row of data in column A
    
    'Set Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    
    startingrow = 2
    'define starting row after headers

    For i = 2 To LastRow
        ' loops through the rows
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'looking forward one cell at next ticker to determine if they match
            
            TotalVol = TotalVol + Cells(i, 7).Value
            'setting total volume to increase
            
            UniqueTicker = Cells(i, 1).Value
            'Get value of new ticker
            
            Opening = Cells(i, 3).Value
            'assign the value of the current row's opening column value
            
            If Cells(startingrow, 3) = 0 Then
            ' handle div/0 error
                Range("I" & SummaryTable).Value = UniqueTicker
                Range("J" & SummaryTable).Value = 0
                Range("K" & SummaryTable).Value = "%" & 0
                Range("L" & SummaryTable).Value = 0
            
            Else
                YearlyChange = Cells(i, 6) - Cells(startingrow, 3)
                'calculate yearly change subtracting earliest opening value from latest closing value for ticker
                
                PercentChange = ((YearlyChange) / Cells(startingrow, 3))
                'calculate percentage change
                Range("I" & SummaryTable).Value = UniqueTicker
                Range("J" & SummaryTable).Value = YearlyChange
                Range("K" & SummaryTable).Value = "%" & (PercentChange * 100)
                Range("L" & SummaryTable).Value = TotalVol

                SummaryTable = SummaryTable + 1
                ' move to next line in summary table
                TotalVol = 0
                ' reset total volume to 0 for next ticker
                startingrow = i + 1
            End If
        Else
            TotalVol = TotalVol + Cells(i, 7).Value
            'continue summing volume while the ticker symbol is the same
        End If

    Next i
    For j = 2 To LastRow
        If Cells(j, 10) > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    TotalVol = 0
    YearlyChange = 0
    PercentChange = 0
    UniqueTicker = ()



End Sub
