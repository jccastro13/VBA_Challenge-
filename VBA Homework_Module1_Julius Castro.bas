Attribute VB_Name = "Module1"
Sub Stock_Ticker():

Dim ws As Worksheet

For Each ws In Worksheets

Dim lastrow As Long
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long
Dim c As Integer
Dim total_volume As Double
total_volume = 0

ws.Range("I1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "% Change"
ws.Range("l1").Value = "Total Stock Volume"

Dim ticker As String
ticker = " "

Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
percent_change = 0
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0

Dim index As Integer
index = 1
open_price = ws.Cells(2, 3).Value
For i = 2 To lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    index = index + 1
    ticker = ws.Cells(i, 1).Value
    ws.Cells(index, "I").Value = ticker
    
    total_volume = total_volume + ws.Cells(i, "G").Value
    ws.Cells(index, "L").Value = total_volume
 
 
    
    total_volume = 0
    close_price = ws.Cells(i, 6).Value
    yearly_change = close_price - open_price
    ws.Cells(index, "j").Value = yearly_change


    If open_price <> 0 Then

    percent_change = (yearly_change / open_price)
    Else
    percent_change = 0
    End If

    ws.Cells(index, "k").Value = percent_change
    open_price = ws.Cells(i + 1, 3).Value
    Else
    total_volume = total_volume + ws.Cells(i, "G").Value
     

End If
   
Next i

Next ws

End Sub


