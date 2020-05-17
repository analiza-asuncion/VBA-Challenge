Attribute VB_Name = "Module1"
Sub Calculate_Stock_Stats():

'ticker symbol
Dim ticker As String

'# of tickers per ws
Dim no_ticker As Integer

'last per ws
Dim lastRow As Long

'opening price per year
Dim op_price As Double

'closing price per year
Dim clo_pric As Double

'yearly change
Dim yr_chg As Double

'% change
Dim pct_change As Double

'totalstockvolume
Dim totalstockvol As Double

'grtst_pct_inc value annually
Dim grtst_pct_inc As Double

'the ticker that has the grtst_pct_inc.
Dim grtst_pct_inc_ticker As String

' grtst_pct_dec
Dim grtst_pct_dec As Double

'grtst_pct_dec
Dim grtst_pct_dec_ticker As String

'grtst stock vol
Dim grtst_stk_vol As Double

'grtst stock vol.
Dim grtst_stk_vol_ticker As String

' each ws loop
For Each ws In Worksheets

    ' initializa
    ws.Activate

    ' Lastrow find
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns per ws
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Init var per ws.
    no_ticker = 0
    ticker = ""
    yr_chg = 0
    op_price = 0
    pct_change = 0
    totalstockvol = 0
    
    ' Skip first row
    For i = 2 To lastRow

        ' current val ticker symbol
        ticker = Cells(i, 1).Value
        
        ' opening price 
        If op_price = 0 Then
            op_price = Cells(i, 3).Value
        End If
        
        ' total stock vol
        totalstockvol = totalstockvol + Cells(i, 7).Value
        
        ' next diff ticker
        If Cells(i + 1, 1).Value <> ticker Then
            ' Increment the number of tickers when we get to a different ticker in the list.
            no_ticker = no_ticker + 1
            Cells(no_ticker + 1, 9) = ticker
            
            clo_pric = Cells(i, 6)

            yr_chg = clo_pric - op_price

            Cells(no_ticker + 1, 10).Value = yr_chg
            
            If yr_chg > 0 Then
                Cells(no_ticker + 1, 10).Interior.ColorIndex = 4

            ElseIf yr_chg < 0 Then
                Cells(no_ticker + 1, 10).Interior.ColorIndex = 3

            Else
                Cells(no_ticker + 1, 10).Interior.ColorIndex = 6
            End If
            

            If op_price = 0 Then
                pct_change = 0
            Else
                pct_change = (yr_chg / op_price)
            End If
            
            Cells(no_ticker + 1, 11).Value = Format(pct_change, "Percent")

            op_price = 0

            Cells(no_ticker + 1, 12).Value = totalstockvol

            totalstockvol = 0
        End If
        
    Next i
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    grtst_pct_inc = Cells(2, 11).Value
    grtst_pct_inc_ticker = Cells(2, 9).Value
    grtst_pct_dec = Cells(2, 11).Value
    grtst_pct_dec_ticker = Cells(2, 9).Value
    grtst_stk_vol = Cells(2, 12).Value
    grtst_stk_vol_ticker = Cells(2, 9).Value

    For i = 2 To lastRow
    
        If Cells(i, 11).Value > grtst_pct_inc Then
            grtst_pct_inc = Cells(i, 11).Value
            grtst_pct_inc_ticker = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value < grtst_pct_dec Then
            grtst_pct_dec = Cells(i, 11).Value
            grtst_pct_dec_ticker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > grtst_stk_vol Then
            grtst_stk_vol = Cells(i, 12).Value
            grtst_stk_vol_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    Range("P2").Value = Format(grtst_pct_inc_ticker, "Percent")
    Range("Q2").Value = Format(grtst_pct_inc, "Percent")
    Range("P3").Value = Format(grtst_pct_dec_ticker, "Percent")
    Range("Q3").Value = Format(grtst_pct_dec, "Percent")
    Range("P4").Value = grtst_stk_vol_ticker
    Range("Q4").Value = grtst_stk_vol
    
Next ws
End Sub

