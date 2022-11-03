# VBA-trial-coursework


'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.


'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.


'The total stock volume of the stock.

'Dim yearlychange As Double
'Dim percentagechange As Double
'Dim totalstockvolume As Integer
'Dim ws As Worksheet

Sub ticker()


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

Next ws

Dim a As Long
a = 1

Dim ticker As String
ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

Dim last As Long
Dim i As Long
Dim j As Integer

last = ws.Cells(Rows.Count, 1).End(xlDown)

Dim openprice As Double
openprice = 0
Dim closeprice As Double
closeprice = 0
Dim pricechange As Double
pricechange = 0
Dim change_percent As Double
change_percent = 0


For i = 2 To last

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        a = a + 1
        ticker = ws.Cells(i, 1).Value
        ws.Cells(a, "I").Value = ticker
    
    closeprice = ws.Cells(i, 6).Value
    change_percent = close_price - open_price
    
    End If
    
    If open_price <> 0 Then
    change_percentage = (change_percentage / open_price) * 100
    
    End If
    

Next i

End Sub
