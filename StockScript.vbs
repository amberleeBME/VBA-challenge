Sub Stock():
' Create a script that will loop through all the stocks for one year and output the following Information:
'   * The ticker symbol
'   * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The total stock volume of the stock.
'-------------------------------------------------
    
    'Define variables
    Dim tick As String
    Dim yDiff As Double
    Dim pDiff As Double
    Dim totalVol as Long
    
    
    'Print column headers
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
        
End Sub