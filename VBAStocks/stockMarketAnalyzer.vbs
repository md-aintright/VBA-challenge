Sub stockMarketAnalyzer()
    ' Loop through all sheets
    For Each ws In Worksheets

        ' Insert new column headers (ticker, yearly change, percent change, total stock volume)
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"

        ' Keep track of results row and open, close stock price
        Dim currentRow As Integer
        currentRow = 2
        Dim closingPrice As Double
        Dim openPrice As Double
        Dim volume As Double

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each stock row
        For i = 2 To lastRow
        
            'If ticker value in previous row is different than current row, Find opening value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Assign open of year value
                openPrice = ws.Cells(i, 3)
                volume = volume + ws.Cells(i, 7).Value

            ' Searches for when ticker value in next row is different than that of the current cell
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker value
                ws.Range("I" & currentRow).Value = ws.Cells(i, 1)
                ' Assign close of year value
                closingPrice = ws.Cells(i, 6)
                volume = volume + ws.Cells(i, 7).Value
                ' Calculate yearly change
                ws.Range("J" & currentRow).Value = (closingPrice - openPrice)

                ' Format yearly change based on positive/negative value
                If ws.Range("J" & currentRow).Value > 0 Then
                    ws.Range("J" & currentRow).Interior.ColorIndex = 4
                Elseif ws.Range("J" & currentRow).Value < 0 Then
                    ws.Range("J" & currentRow).Interior.ColorIndex = 3
                End if

                ' Calculate change percentage
                If openPrice = 0 Then ' Check if divisor is zero, if so set to zero 
                    ws.Range("K" & currentRow).Value = 0
                Else
                    ws.Range("K" & currentRow).Value = ((closingPrice - openPrice) / openPrice)
                End If

                ' Format change percentage cell as percentage
                ws.Range("K" & currentRow).NumberFormat = "0.00%"
                ' Set total stock volume
                ws.Range("L" & currentRow).Value = volume
                ' Move to next result row
                currentRow = currentRow + 1
                ' Reset volume, open, close
                openPrice = 0
                closingPrice = 0
                volume = 0
            Else
            ' Add volume to total volume
            volume = volume + ws.Cells(i, 7).Value
            End If
            ' Next row
        Next i

        ' CHALLENGE

        'Add title to challenge row cells
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        ws.Range("N2") = "Greatest % Increase"
        ws.Range("N3") = "Greatest % Decrease"
        ws.Range("N4") = "Greatest Total Volume"

        ' Determine the last row of results
        lastRowResults = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ' Create/set variables for max/min and to track row number
        Dim maxRow As Integer
        Dim minRow As Integer
        Dim maxVolumeRow As Integer
        Dim maxPercentage As Double
        maxPercentage = 0
        Dim minPercentage As Double
        minPercentage = 0
        Dim maxVolume As Double
        maxVolume = 0

        ' Loop through each row of results
        For j = 2 To lastRowResults

            ' Find MAX percentage, increase row number if found
            If ws.Range("K" & j).Value > maxPercentage Then
                maxPercentage = ws.Range("K" & j).Value
                maxRow = j
            End If

            ' Find MIN percentage, increase row number if found
            If ws.Range("K" & j).Value < minPercentage Then
                minPercentage = ws.Range("K" & j).Value
                minRow = j
            End If

            ' Find MAX volume, increase row number if found
            If ws.Range("L" & j).Value > maxVolume Then
                maxVolume = ws.Range("L" & j).Value
                maxVolumeRow = j
            End If

        Next j

        'Set ticker value based on row variables
        ws.Range("O2") = ws.Range("I" & maxRow).Value
        ws.Range("O3") = ws.Range("I" & minRow).Value
        ws.Range("O4") = ws.Range("I" & maxVolumeRow).Value
        ' Set Max percentage
        ws.Range("P2") = maxPercentage
         ' Format change percentage cell as percentage
        ws.Range("P2").NumberFormat = "0.00%"
        ' Set Min percentage
        ws.Range("P3") = minPercentage
        ' Format change percentage cell as percentage
        ws.Range("P3").NumberFormat = "0.00%"
        ' Set max volume
        ws.Range("P4") = maxVolume

    Next ws  

End Sub