Sub WallStreet()

    Dim ticker As String

    Dim current_ticker As String

    Dim volume As Double

    Dim open_stock as Double

    Dim close_stock as Double

    Dim output As Integer

    Dim ws As Worksheet

    Dim greatest As Double

    For Each ws In Worksheets

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        open_stock = 0
        close_stock = 0
        volume = 0
        output = 2

        For i = 2 To lastrow
        

            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

                open_stock = ws.Cells(i, 3).Value
                volume = volume + ws.Cells(i, 7).Value
                
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value then   
                
                close_stock = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("I" & output).Value = ticker
                ws.Range("L" & output).Value = volume
                ws.Range("J" & output).Value = close_stock - open_stock

                If ws.Range("J" & output).Value > 0 Then

                   ws.Range("J" & output).Interior.ColorIndex = 4
                
                Else
                    
                   ws.Range("J" & output).Interior.ColorIndex = 3
                
                End IF

                If open_stock = 0 then

                    ws.Range("K" & output).Value = 1
                
                Else

                    ws.Range("K" & output).Value = (close_stock - open_stock) / open_stock

                End If
                
                ws.Range("K" & output).NumberFormat = "0.00%"
                volume = 0
                output = output + 1

            Else

                volume = volume + ws.Cells(i, 7).Value
        
            End If
        
        
        Next i
        
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker" 
        ws.Cells(1, 17).Value = "Value"
        greatest = 0

        for j = 2 to lastrow

            current_ticker = ws.Cells(j, 9).Value

            If ws.Cells(j, 11).Value > greatest Then

                greatest = ws.Cells(j, 11).Value
                ws.Range("Q2").Value = greatest
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = current_ticker

            End If

        next j

        greatest = 0
        
        for k = 2 to lastrow

            current_ticker = ws.Cells(k, 9).Value

            If ws.Cells(k, 11).Value < greatest Then

                greatest = ws.Cells(k, 11).Value
                ws.Range("Q3").Value = greatest
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = current_ticker

            End If

        next k

        greatest = 0

        for l = 2 to lastrow

            current_ticker = ws.Cells(l, 9).Value

            If ws.Cells(l, 12).Value > greatest Then

                greatest = ws.Cells(l, 12).Value
                ws.Range("Q4").Value = greatest
                ws.Range("P4").Value = current_ticker

            End IF

        next l

        ws.Cells.Columns.AutoFit

    Next ws

End Sub