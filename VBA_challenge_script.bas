Attribute VB_Name = "Module1"
Sub Hw1Script()

Dim ws As Worksheet

For Each ws In Worksheets
    
    Dim ticker As String
    Dim vol As Double
    Dim table_row As Integer
    Dim price_open As Double
    Dim price_close As Double
    Dim yearly_price_change As Double
    Dim yearly_percent_change As Double

    ticker = " "
    vol = 0
    price_open = 0
    price_close = 0
    yearly_price_change = 0
    yearly_percent_change = 0
    table_row = 2

    Dim r As Integer
    r = ws.Range("A1").End(xlToRight).Column + 2
    ws.Cells(1, r).Value = "Ticker"
    ws.Cells(1, r + 1).Value = "Yearly Price Change"
    ws.Cells(1, r + 2).Value = "Yearly Percentage Change"
    ws.Cells(1, r + 3).Value = "Total Volume"
    
   
    Dim findrng As Range
    Dim index1 As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim index4 As Long
    Set findrng = ws.Range(Cells(1, 1).Address, Cells(1, 8).Address)
    index1 = findrng.Find("<ticker>").Column
    index2 = findrng.Find("<open>").Column
    index3 = findrng.Find("<close>").Column
    index4 = findrng.Find("<vol>").Column
    

    price_open = ws.Cells(2, index2).Value
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 3 To lastrow
        If ws.Cells(i + 1, index1).Value <> ws.Cells(i, index1).Value Then
            ticker = ws.Cells(i, index1).Value
            price_close = ws.Cells(i, index3).Value
            yearly_price_change = price_close - price_open
        
        If price_open <> 0 Then
            yearly_percent_change = (yearly_price_change / price_open) * 100
        End If
         
        vol = vol + ws.Cells(i, index4).Value
        
        ws.Cells(table_row, r).Value = ticker
        ws.Cells(table_row, r + 1).Value = yearly_price_change
        ws.Cells(table_row, r + 2).Value = (CStr(yearly_percent_change) & "%")
        ws.Cells(table_row, r + 3).Value = vol
        
        If (yearly_price_change > 0) Then
            ws.Cells(table_row, r + 1).Interior.ColorIndex = 4
            ElseIf (yealy_price_change <= 0) Then
                ws.Cells(table_row, r + 1).Interior.ColorIndex = 3
                End If
                
        table_row = table_row + 1
        
        price_open = ws.Cells(i + 1, index2).Value
        yearly_percent_change = 0
        vol = 0
        
        Else
            vol = vol + ws.Cells(i, index4).Value
        End If
            
    Next i
Next ws

End Sub
