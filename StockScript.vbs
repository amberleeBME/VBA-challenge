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
    Dim totalVol As Long
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    'For Each ws In Worksheets

        ' --------------------------------------------
        ' SETUP OUTPUT TABLES
        ' --------------------------------------------
        
        'Output Table 1: Column Headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Output Table 2: Row Headers
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Output Table 2: Column Headers
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' --------------------------------------------
        ' INITIALIZE VARIABLES AND DETERMINE LAST ROW
        ' --------------------------------------------
        
        'Initialize variables using values from first record (values from row '2' of current spreadsheet)
        tick = Range("A2").Value
        yDiff = Range("C2").Value
        pDiff = Range("C2").Value
        totalVol = Range("G2").Value
        
        'Determine the last row in the worksheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' --------------------------------------------
        ' ITERATE THROUGH EACH ROW
        ' --------------------------------------------
        'For i = 2 To LastRow
            
        'Next i
    'Next ws
End Sub