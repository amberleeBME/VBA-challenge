' --------------------------------------------
' INSTRUCTIONS
' --------------------------------------------
' 1. Create a script that will loop through all the stocks for one year and output the following:
'           * The ticker symbol
'           * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'           * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'           * The total stock volume of the stock.
' 2. Have the script apply conditional formatting that will highlight positive change in green and negative change in red.
' -------
' *Bonus*
' -------
' 3. The script should also output:
'           * The stock with the Greatest % increase
'           * The stock with the Greatest % decrease
'           * The stock with the Greatest total volume
' 4. Running the VBA script once should generate an output on every worksheet (every year).

Sub Stock():
    ' --------------------------------------------
    ' DEFINE VARIABLES
    ' --------------------------------------------
    Dim tick As String
    Dim yOpen As Double
    Dim pDiff As Double
    Dim totalVol As Double
    Dim outRow As Integer

    Dim minP As Double
    Dim maxP As Double
    Dim maxV As Double
    Dim tickMinP As String
    Dim tickMaxP As String
    Dim tickMaxV As String
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        ' Set current worksheet
        Set curr_ws = Worksheets(ws.Name)
        ' --------------------------------------------
        ' SETUP TABLE HEADERS
        ' --------------------------------------------
        ' Output Table 1: Column Headers
        curr_ws.Range("I1").Value = "Ticker"
        curr_ws.Range("J1").Value = "Yearly Change"
        curr_ws.Range("K1").Value = "Percent Change"
        curr_ws.Range("L1").Value = "Total Stock Volume"
        
        ' Output Table 2: Row Headers
        curr_ws.Range("O2").Value = "Greatest % Increase"
        curr_ws.Range("O3").Value = "Greatest % Decrease"
        curr_ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Output Table 2: Column Headers
        curr_ws.Range("P1").Value = "Ticker"
        curr_ws.Range("Q1").Value = "Value"
        
        ' --------------------------------------------
        ' INITIALIZE VARIABLES AND DETERMINE LAST ROW
        ' --------------------------------------------
        yOpen = curr_ws.Range("C2").Value
        totalVol = curr_ws.Range("G2").Value
        outRow = 2
        minP = 0
        maxP = 0
        maxV = 0
        
        'Determine the last row in the worksheet
        LastRow = curr_ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' --------------------------------------------
        ' LOOP THROUGH ALL ROWS
        ' --------------------------------------------
        For i = 2 To LastRow
            tick = curr_ws.Cells(i, 1).Value                ' Current ticker
            newTick = curr_ws.Cells(i + 1, 1).Value         ' New ticker name
            newOpen = curr_ws.Cells(i + 1, 3).Value         ' New opening price
            yClose = curr_ws.Cells(i + 1, 6).Value          ' New closing price
            newVol = curr_ws.Cells(i + 1, 7).Value          ' New volume
            
            ' If the new ticker names are the same, then add to total, else add to Output Table 1
            If tick = newTick Then
                totalVol = totalVol + newVol
            Else
                
                ' Find yearly change.
                yDiff = curr_ws.Cells(i, 6).Value - yOpen
                
                ' Calculate percent change. If opening price = 0, the percent change = yearly change.
                If yOpen <> 0 Then
                    pDiff = yDiff / yOpen
                Else
                    pDiff = yDiff
                End If
                ' --------------------------------------------
                ' OUTPUT TABLE 1
                ' --------------------------------------------
                ' Enter values onto Output Table 1
                curr_ws.Cells(outRow, 9).Value = tick
                curr_ws.Cells(outRow, 10).Value = yDiff
                curr_ws.Cells(outRow, 11).Value = pDiff
                curr_ws.Cells(outRow, 12).Value = totalVol
                ' -------
                ' *Format Table 1*
                ' -------
                ' If yearly change is negative, then red
                ' Else if yearly change is positive, then green
                If yDiff < 0 Then
                    curr_ws.Cells(outRow, 10).Interior.ColorIndex = 3
                ElseIf yDiff > 0 Then
                    curr_ws.Cells(outRow, 10).Interior.ColorIndex = 4
                Else
                End If
                
                ' Format Percent Change column
                curr_ws.Cells(outRow, 11).NumberFormat = "0.00%"
                
                ' Determine if greatest percent increase, percent decrease, or total volume
                If pDiff > maxP Then
                    maxP = pDiff
                    tickMaxP = tick
                ElseIf pDiff < minP Then
                    minP = pDiff
                    tickMinP = tick
                Else
                End If
                If totalVol > maxV Then
                    maxV = totalVol
                    tickMaxV = tick
                End If
                
                ' --------------------------------------------
                ' RESET VARIABLES FOR NEW STOCK
                ' --------------------------------------------
                totalVol = newVol
                yOpen = newOpen
                
                ' next row on Output Table 1
                outRow = outRow + 1
            End If
        Next i
        ' --------------------------------------------
        ' OUTPUT TABLE 2
        ' --------------------------------------------
        curr_ws.Range("P2").Value = tickMaxP
        curr_ws.Range("Q2").Value = maxP
        curr_ws.Range("P3").Value = tickMinP
        curr_ws.Range("Q3").Value = minP
        curr_ws.Range("P4").Value = tickMaxV
        curr_ws.Range("Q4").Value = maxV
        'Format percent cells
        curr_ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ' Autofit to display data
        curr_ws.Columns("I:Q").AutoFit
    Next ws
End Sub