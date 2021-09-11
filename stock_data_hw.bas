Attribute VB_Name = "Module1"
Sub wallStreet():

    'Define a variable to hold the ticker name
    Dim ticker As String

    'Define variables to hold the opening price, closing price, yearly change in price, and percent change
    Dim open_price As Double
    Dim close_price As Double
    Dim year_change As Double
    Dim per_change As Double

    'Define variables to hold the sum of the stock volume for the year and the daily volume
    Dim total_volume As Double
    Dim day_volume As Double

    'Keep track of the row of the summary table
    Dim sum_table_row As Double
    
    'Define variable to help find the opening price
    Dim start As Double

    'Loop through each worksheet in the workbook
    For Each ws In Worksheets
        
        'Label the summary table header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Reset the start variable, total volume, daily volume, summary table row at the beginning of each worksheet
        start = 2
        total_volume = 0
        sum_table_row = 2

        'Find the last row of data (this command is from VBA class #2, the first wells fargo activity)
        Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through stock data
        For i = 2 To LastRow
            ticker = ws.Cells(i, 1).Value
            next_ticker = ws.Cells(i + 1, 1).Value

            'Set the daily volume
            day_volume = ws.Cells(i, 7).Value

            'Check to see if we are on the last row of that ticker
            If next_ticker <> ticker Then

               'Add to the total volume
                total_volume = total_volume + day_volume

                'Set the first non-zero opening price for the current ticker (code based on help received from Learning Assistant on Slack)
                If ws.Cells(start, 3).Value = 0 Then
                    For finder = start To i
                        If ws.Cells(finder, 3).Value <> 0 Then
                            start = finder
                            Exit For
                        End If
                    Next finder
                End If
                open_price = ws.Cells(start, 3).Value
                
               'Deal with total volume and/or open prices that are 0 for a ticker
                If (total_volume = 0 Or open_price = 0) Then
                    ws.Cells(sum_table_row, 9).Value = ticker
                    ws.Cells(sum_table_row, 10).Value = 0
                    ws.Cells(sum_table_row, 11).Value = 0 & "%"
                    ws.Cells(sum_table_row, 12).Value = total_volume
                    
                    'change the ticker variable to the new ticker, add to the summary table row number
                    sum_table_row = sum_table_row + 1
    
                    'reset the total volume to 0
                    total_volume = 0
    
                    'reset the start variable to help find the next ticker's opening price
                    start = i + 1
                Else
                    'Set the closing price for the current ticker, calculate the yearly change, add to the total volume,
                    close_price = ws.Cells(i, 6).Value
                    year_change = close_price - open_price
                    per_change = Round((year_change / open_price * 100), 2)
    
                    'place the results in the summary table
                    ws.Cells(sum_table_row, 9).Value = ticker
                    ws.Cells(sum_table_row, 10).Value = year_change
                    ws.Cells(sum_table_row, 11).Value = per_change & "%"
                    ws.Cells(sum_table_row, 12).Value = total_volume
                    
                    'change the color of the percent change cells green if greater than 0, red if less than 0
                    If ws.Cells(sum_table_row, 10).Value < 0 Then
                        ws.Cells(sum_table_row, 10).Interior.ColorIndex = 3
                    ElseIf ws.Cells(sum_table_row, 10).Value > 0 Then
                        ws.Cells(sum_table_row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(sum_table_row, 10).Interior.ColorIndex = 0
                    End If
    
                    'change the ticker variable to the new ticker, add to the summary table row number
                    sum_table_row = sum_table_row + 1
    
                    'reset the total volume to 0
                    total_volume = 0
    
                    'reset the start variable to help find the next ticker's opening price
                    start = i + 1
                End If
            Else
                'If not - add the daily volume to the total volume, reset the daily volume to 0
                total_volume = total_volume + day_volume

            End If

        Next i

    Next ws

End Sub



