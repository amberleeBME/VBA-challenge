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
' *BONUS*
' -------
' 3. The script should also output:
'           * The stock with the Greatest % increase
'           * The stock with the Greatest % decrease
'           * The stock with the Greatest total volume
' 4. Running the VBA script once should generate an output on every worksheet (every year).

Sub Stock():

    ' --------------------------------------------
    ' VARIABLES
    ' --------------------------------------------
    
    'Define variables
    Dim tick As String
    Dim yDiff As Double
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
        ' SETUP OUTPUT TABLES
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
        ' INITIALIZE VARIABLES AND DETERMINE LAST ROW
        ' --------------------------------------------
        
        'Initialize variables using values from first record (values from row '2' of current spreadsheet)
        outRow = 2
        totalVol = curr_ws.Range("G2").Value
        
        'Determine the last row in the worksheet
        LastRow = curr_ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' --------------------------------------------
        ' LOOP THROUGH EACH ROW
        ' --------------------------------------------
        For i = 2 To LastRow
            tick = curr_ws.Cells(i, 1).Value
            newTick = curr_ws.Cells(i + 1, 1).Value
            newVol = curr_ws.Cells(i + 1, 7).Value
            If tick <> newTick Then
                curr_ws.Cells(outRow, 9).Value = tick
                curr_ws.Cells(outRow, 12).Value = totalVol
                tick = newTick
                totalVol = newVol
                outRow = outRow + 1
            Else
                totalVol = totalVol + newVol
            End If
        Next i
    Next ws
End Sub