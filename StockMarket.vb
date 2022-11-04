Sub StockMarket()
'Create a script that loops through all the stocks for one year and outputs the following information:
For Each ws In Worksheets

    'create a column header for each section
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'change column width
    ws.Columns("I:L").AutoFit
    '* The ticker symbol.
    Dim Ticker_Symbol As String
    'adding a row to the summary table
    Dim Table_Row As Long
    Table_Row = 2
    'variable for total volume of each ticker
    Dim Total_Volume As Double
    Total_Volume = 0
    'variable to calculate yearly change
    Dim Yearly_Change As Double
    Yearly_Change = 0
    'varibale to hold stock at the beginning of the year
    Dim Stock_Open As Double
    Stock_Open = 0
    'variable to hold stock at the end of the year
    Dim Stock_Close As Double
    Stock_Close = 0
    'variable to calculate percent change
    Dim Percent_Change As Double
    Percent_Change = 0
    'find last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'print the ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
            ws.Range("I" & Table_Row).Value = Ticker_Symbol
            'total stock volume of the stock and print
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Table_Row).Value = Total_Volume
            '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year. print
            Stock_Close = ws.Cells(i, 6).Value
            Yearly_Change = Stock_Close - Stock_Open
            ws.Range("J" & Table_Row).Value = Yearly_Change
            'Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
            If Yearly_Change < 0 Then
                ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                End If
            'percent change = final value-initial value / final value
            Percent_Change = (Stock_Close - Stock_Open) / Stock_Open
            ws.Range("K" & Table_Row).Value = Percent_Change
            'format percent_change as a percentage
            ws.Range("K" & Table_Row).NumberFormat = "0.00%"
            'add one to the summary row
            Table_Row = Table_Row + 1
            'reset the volume, yearly change, and percent change for next ticker
            Total_Volume = 0
            Yearly_Change = 0
            Stock_Open = 0
            Stock_Close = 0
            Percent_Change = 0
        Else
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Stock_Open = ws.Cells(i, 3).Value
            End If
        End If

        Next i
    Next ws
End Sub
