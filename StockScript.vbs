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
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        ' Set current worksheet
        Set curr_ws = Worksheets(ws.Name)
        
        ' --------------------------------------------
        ' SETUP TABLE HEADERS
        ' --------------------------------------------
        
        'Output Table 1: Column Headers
        curr_ws.Range("I1").Value = "Ticker"
        curr_ws.Range("J1").Value = "Yearly Change"
        curr_ws.Range("K1").Value = "Percent Change"
        curr_ws.Range("L1").Value = "Total Stock Volume"
        
        'Output Table 2: Row Headers
        curr_ws.Range("O2").Value = "Greatest % Increase"
        curr_ws.Range("O3").Value = "Greatest % Decrease"
        curr_ws.Range("O4").Value = "Greatest Total Volume"
        
        'Output Table 2: Column Headers
        curr_ws.Range("P1").Value = "Ticker"
        curr_ws.Range("Q1").Value = "Value"
        
        ' Autofit to display data
        curr_ws.Columns("I:Q").AutoFit
        
        ' --------------------------------------------
        ' INITIALIZE AND DETERMINE LAST ROW
        ' --------------------------------------------
        yOpen = curr_ws.Range("C2").Value
        totalVol = curr_ws.Range("G2").Value
        outRow = 2
        
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
            
            ' If the new ticker names are the same, then add to total
            If tick = newTick Then
                totalVol = totalVol + newVol
            Else
                ' Find yearly change
                yDiff = curr_ws.Cells(i, 6).Value - yOpen
                
                ' If necessary, change opening price to 1 to avoid dividing by 0 errror.
                If yOpen = 0 Then
                    yOpen = 1
                End If
                
                ' Find percent change
                pDiff = yDiff / yOpen
                
                ' Output current ticker name, yearly change, percent change, and total volume
                curr_ws.Cells(outRow, 9).Value = tick
                curr_ws.Cells(outRow, 10).Value = yDiff
                curr_ws.Cells(outRow, 11).Value = pDiff
                curr_ws.Cells(outRow, 12).Value = totalVol
                
                ' Reset total and opening price for new ticker
                totalVol = newVol
                yOpen = newOpen
                
                ' --------------------------------------------
                ' FORMAT CELLS
                ' --------------------------------------------
                ' *Color Formatting*
                ' -------
                ' If yearly change is negative, then red
                ' Else if yearly change is positive, then green
                If yDiff < 0 Then
                    curr_ws.Cells(outRow, 10).Interior.ColorIndex = 3
                ElseIf yDiff > 0 Then
                    curr_ws.Cells(outRow, 10).Interior.ColorIndex = 4
                Else
                End If
                
                ' -------
                ' *Number Formatting*
                ' -------
                ' If yearly change is negative, then red
                curr_ws.Cells(outRow, 11).NumberFormat = "0.00%"
                ' next output row
                outRow = outRow + 1
                
            End If
        Next i
    Next ws
End Sub