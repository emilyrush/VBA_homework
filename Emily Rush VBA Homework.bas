Attribute VB_Name = "Module1"
Sub tickerTickerTiming()
    'loop through all the sheets
    For ws = 1 To Worksheets.Count
    
        'Name the summary table columns
        Cells(1, 9).Value = "Stock Symbol"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Set variables for new column items
        Dim stock_name, next_name, current_name As String
        Dim year_open, year_close, open1, close1 As Double
    
        Dim percent_change As Variant
        Dim stock_total As Variant
        Dim yearly_change As Variant
    
        stock_total = 0
        yearly_change = 0
        percent_change = 0
        
        'Set a variable for the last row
        Dim lastRow As Long
        lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
        
        'Keep track of the location of the stock in the summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
    
        'Loop through stock names
        For i = 2 To lastRow
        
        'initialize stock name
        If stock_name = "" Then
            stock_name = Cells(i, 1).Value
            year_open = Cells(i, 3).Value
        End If
        
        current_name = Cells(i, 1).Value
        next_name = Cells(i + 1, 1).Value
        stock_total = stock_total + Cells(i, 7).Value
        
        If next_name <> current_name Then
            'value of stock at the end of the year
            year_close = Cells(i, 6).Value
             
             'Print stock name to the summary table
             Range("I" & summary_table_row).Value = stock_name
             
             'Add the brand's $$$ to the summary table
             Range("L" & summary_table_row).Value = stock_total
            
            'Calculate the yearly and percent changes and add them to the summary table
            yearly_change = year_close - year_open
            Range("J" & summary_table_row).Value = yearly_change
            If year_open <> 0 Then
                percent_change = yearly_change / year_open
                Range("K" & summary_table_row).Value = percent_change
            End If
    
            
             'In the summary table, move down one row.
             summary_table_row = summary_table_row + 1
             
             'and reset the stock total
             stock_total = 0
             yearly_change = 0
             stock_name = Cells(i + 1, 1).Value
             year_open = Cells(i + 1, 3).Value
        End If
        Next i
        
        'Make pretty
        ActiveSheet.Columns("I:L").AutoFit
        ActiveSheet.Columns("K:K").NumberFormat = "0.00%"
        
        'Give it some color
        For j = 2 To lastRow
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        'Check to see if this is the last sheet. If not, move on to the next.
        If ActiveSheet.Index <> ActiveWorkbook.Worksheets.Count Then
            Worksheets(ActiveSheet.Index + 1).Select
        End If

    Next ws
    
    MsgBox ("AWWWW YEAAAA")
End Sub
