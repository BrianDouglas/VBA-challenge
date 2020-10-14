Attribute VB_Name = "Module1"
Sub summarize()

    Dim data_length, output_row As Integer
    Dim newyear_opening, endyear_closing, change, change_percent, saved_percents(2) As Double
    Dim total_volume, saved_total_volume As LongLong
    Dim current_ticker, next_ticker, saved_tickers(3) As String
    
    For n = 1 To Sheets.Count
        Sheets(n).Select
        'reset cumulative variables
        Erase saved_percents
        Erase saved_tickers
        saved_total_volume = 0
        'Configure output columns 1
        output_row = 1
        Cells(output_row, 9).Value = "Ticker"
        Cells(output_row, 10).Value = "Yearly Change"
        Cells(output_row, 11).Value = "Percent Change"
        Columns(11).NumberFormat = "0.00%"
        Cells(output_row, 12).Value = "Total Stock Volume"
        Range(Columns(9), Columns(12)).AutoFit
        output_row = output_row + 1
        
        'Configure output columns 2
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Columns(14).AutoFit
        
        'get length of data
        data_length = Cells(Rows.Count, 1).End(xlUp).Row
        
        'ensure total_volume is initialized to 0
        total_volume = 0
        'get opening value for first stock before looping
        newyear_opening = Cells(2, 3).Value
        
        'loop through data
        For i = 2 To data_length
            'increse total volume from new row
            total_volume = total_volume + Cells(i, 7).Value
            'update values for ticker symbol pointers
            current_ticker = Cells(i, 1).Value
            next_ticker = Cells(i + 1, 1).Value
            'check if we're at a new symbol
            If current_ticker <> next_ticker Then
                'get closing data and calculate change
                endyear_closing = Cells(i, 6).Value
                change = endyear_closing - newyear_opening
                If newyear_opening = 0 Then
                    change_percent = 0
                Else
                    change_percent = change / newyear_opening
                End If
                'update saved variables
                If change_percent > saved_percents(0) Then
                    saved_percents(0) = change_percent
                    saved_tickers(0) = current_ticker
                End If
                If change_percent < saved_percents(1) Then
                    saved_percents(1) = change_percent
                    saved_tickers(1) = current_ticker
                End If
                If total_volume > saved_total_volume Then
                    saved_total_volume = total_volume
                    saved_tickers(2) = current_ticker
                End If
                'output data
                Cells(output_row, 9).Value = current_ticker
                Cells(output_row, 10).Value = change
                'change cell color based on change
                If change >= 0 Then
                    Cells(output_row, 10).Interior.ColorIndex = 4
                Else
                    Cells(output_row, 10).Interior.ColorIndex = 3
                End If
                
                Cells(output_row, 11).Value = change_percent
                Cells(output_row, 12).Value = total_volume
                output_row = output_row + 1
                
                'reset variables for next ticker symbol
                total_volume = 0
                newyear_opening = Cells(i + 1, 3).Value
                            
            End If
            
        Next i
        'update output columns 2
        Cells(2, 15).Value = saved_tickers(0)
        Cells(3, 15).Value = saved_tickers(1)
        Cells(4, 15).Value = saved_tickers(2)
        Cells(2, 16).Value = saved_percents(0)
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(3, 16).Value = saved_percents(1)
        Cells(3, 16).NumberFormat = "0.00%"
        Cells(4, 16).Value = saved_total_volume
    Next n
End Sub
