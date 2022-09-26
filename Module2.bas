Attribute VB_Name = "Module2"
Sub hw_two()
    'defining variables
    Dim opening As Double
    'calculate last row
    Dim last_Row As Double
    last_Row = Cells(Rows.Count, "A").End(xlUp).Row
    Dim closing As Double
    'initial value of counter #2
    Dim x As Double
    x = 2
    Dim ticker_fill_row As Double
    ' Set an initial variable for holding the total per ticker
    Dim ticker_total As Double
    ticker_total = 0
    ' Keep track of the location for each ticker in the summary table
    Dim final_row As Integer
    final_row = 2
    'defining opening value
    opening = Cells(2, 3).Value
    'defining ticker fill row
    ticker_fill_row = 2
    
    'filling header titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To last_Row
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'ticker fill statement
            Cells(ticker_fill_row, 9).Value = Cells(i, 1)
            ticker_fill_row = ticker_fill_row + 1
        End If
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'checking if cells are the same value
            closing = Cells(i, 6).Value ' defining closing value
            Cells(x, 10).Value = closing - opening ' calculating yearly change
            ' Add to the ticker total
            ticker_total = ticker_total + Cells(i, 7).Value
            ' Print the ticker total to the Summary Table
            Range("L" & final_row).Value = ticker_total
            ' Add one to the summary table row
            final_row = final_row + 1
            ' Reset the ticker Total
            ticker_total = 0
                If Cells(x, 10).Value < 0 Then 'conditional formatting of colors
                     Cells(x, 10).Interior.ColorIndex = 3
                Else
                     Cells(x, 10).Interior.ColorIndex = 4
                End If
            Cells(x, 11).Value = ((closing - opening) / opening)
            Range("K" & x).Value = FormatPercent(Range("K" & x)) 'percentage change
            'Resets
            x = x + 1
            opening = Cells(i + 1, 3)
        Else
            ' Add to the ticker Total
            ticker_total = ticker_total + Cells(i, 7).Value
        End If
    Next i
End Sub

