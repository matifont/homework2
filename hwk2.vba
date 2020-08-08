Sub TickerInfo()
Dim ticker As String
Dim total_volume As Double
Dim openclosecounter, ticker_counter As Double
Dim yearly_open, yearly_end As Double

    For Each ws In Worksheets 'Pass the code in every worksheet
        total_volume = 0
        ticker_counter = 2 'row to write out ticker summary
        openclosecounter = 2 'row to save open and close values

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"

        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            'The counter will start in row two and will have one added after it finds a value that is different to the previous cell
            total_volume = total_volume + Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(openclosecounter, 3)

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'output el summary de la variable antes de incrementar el contador
                yearly_end = ws.Cells(i, 6)
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = yearly_end - yearly_open
                'if the opening value is equal to 0 then we will get an error in the percentage or a negative number so we will just put null value
                If yearly_open = 0 Then
                    ws.Cells(ticker_counter, 11).Value = Null
                Else
                    ws.Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
                End If
                ws.Cells(ticker_counter, 12).Value = total_volume

                'Color cell green if it i> 0 and red < 0
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If

                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"

                'Reset counter
                total_volume = 0
                ticker_counter = ticker_counter + 1
                ticker_open_close_counter = i + 1
            End If

         Next i
    Next ws
End Sub
Sub Challenge()
    Call TickerInfo
    For Each ws In Worksheets
        'headers for each colum w/range
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"

        Dim max, min As Double
        Dim min_row_index, max_row_index, max_total_volume_index As Integer
        Dim max_total_volume As Double

        max = 0
        min = 0
        max_total_volume = 0

        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            If ws.Cells(i, 11) > max Then
                max = ws.Cells(i, 11)
                max_row_index = i
            End If

            If ws.Cells(i, 11) < min Then
                min = ws.Cells(i, 11)
                min_row_index = i
            End If

            If ws.Cells(i, 12) > max_total_volume Then
                max_total_volume = ws.Cells(i, 12)
                max_total_volume_index = i
            End If
        Next i
        'write values in the following cells
        ws.Range("P2") = ws.Cells(max_row_index, 9).Value
        ws.Range("P3") = ws.Cells(min_row_index, 9).Value
        ws.Range("P4") = ws.Cells(max_total_volume_index, 9).Value

        ws.Range("Q2") = max
        ws.Range("Q3") = min
        ws.Range("Q4") = max_total_volume

        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    Next ws
End Sub
