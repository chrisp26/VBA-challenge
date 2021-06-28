sub loop_through_stock_sheets_with_challenges

    ' Set initial variables
    Dim ws As Worksheet
    Dim ticker_code as String
    Dim yr_opening_price as Double
    Dim yr_closing_price as Double
    Dim volume as Double
    Dim yr_change as Double
    Dim pc_change as Double
    Dim summary_row as Integer
    Dim last_row as Double
    Dim greatest_pc_increase as Double
    Dim greatest_pc_decrease as Double
    Dim greatest_volume as Double


        ' Loop through all worksheets
        For Each ws In activeworkbook.worksheets
        ws.activate

            'Identify last row
            last_row = ws.cells(Rows.Count, 1).End(xlUp).Row

            'Set number format for %
            ws.range("L2:L" & last_row).numberformat = "0.00%"

            'Set summary row
            summary_row = 2

            'Set headers
            ws.range("J1").value = "Ticker"
            ws.range("k1").value = "yearly change"
            ws.range("l1").value = "percent change"
            ws.range("m1").value = "total stock volume"

            'Set hw 'Challenges' headers
            ws.range("Q1").value = "Ticker"
            ws.range("R1").value = "Value"
            ws.range("P" & summary_row).value = "Greatest % Increase"
            ws.range("P" & summary_row + 1).value = "Greatest % Decrease"
            ws.range("P" & summary_row + 2).value = "Greatest Total Volume"


            'Set initial sheet yr_opening_price
             yr_opening_price = ws.cells(2,3).value
            
            'Set intial 'greatest' values
            greatest_pc_increase = 0
            greatest_pc_decrease = 0
            greatest_volume = 0

                'Loop through all rows and extract data
                For i = 2 to last_row

                    'Check to see if the ticker code has changed
                    If ws.cells(i + 1, 1).value <> ws.cells(i, 1).value Then

                        'If code has changed: publish values to table
                            
                            'Set & print ticker code
                            ticker_code = ws.cells(i,1).value
                            ws.range("J" & summary_row).value = ticker_code

                            'Add to the volume
                            volume = volume + ws.cells(i,7).value

                            'Set closing price
                            yr_closing_price = ws.cells(i, 6).value

                            'Set yearly change and print
                            yr_change = yr_closing_price - yr_opening_price
                                On Error resume Next
                            pc_change = yr_change / yr_opening_price
                                If Err.number <> 0 Then
                                pc_change = 0
                                End if
                            'Alternative: pc_change = (yr_closing_price / yr_opening_price) -1


                            ws.range("k" & summary_row).value = yr_change
                            ws.range("l" & summary_row).value = pc_change

                                'set formatting
                                If pc_change > 0 Then
                                    ws.Range("L" & summary_row).Interior.ColorIndex = 4
                                        
                                ElseIf pc_change < 0 Then
                                    ws.Range("L" & summary_row).Interior.ColorIndex = 3

                                Else 
                                    ws.Range("L" & summary_row).Interior.ColorIndex = 2

                                End If
                            
                            ' Print volume amount in summary table
                            ws.range("m" & summary_row).value = volume
                                
                                ' ------------------------------------------------------------------
                                'Extra challenges
                                    'Set greatest % increase
                                    if pc_change > greatest_pc_increase Then
                                    greatest_pc_increase = pc_change
                                    ws.range("Q2").value = ticker_code
                                    ws.range("R2").value = pc_change
                                    End if

                                    'Set greatest % decrease
                                    if pc_change < greatest_pc_decrease Then
                                    greatest_pc_decrease = pc_change
                                    ws.range("Q3").value = ticker_code
                                    ws.range("R3").value = pc_change
                                    End if

                                     'Set greatest volume
                                    if volume > greatest_volume Then
                                    greatest_volume = volume
                                    ws.range("Q4").value = ticker_code
                                    ws.range("R4").value = volume
                                    End if

                                    'Set number format
                                    ws.range("R2:R3").numberformat = "0.00%"
                                ' ------------------------------------------------------------------

                            'Add one to summary table row
                            summary_row = summary_row + 1

                            'Reset the volume count
                            volume = 0

                            'Set the new year opening price
                                yr_opening_price = ws.cells(i + 1, 3).value
                                
                    Else

                        'If stock code is the same

                            'Add to the volume
                            volume = volume + ws.cells(i,7).value

                    End if

                Next i



        Next ws

msgbox("Your code has finished running. Whoop!")

End sub