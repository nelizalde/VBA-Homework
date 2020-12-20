Attribute VB_Name = "Module1"
Sub Ticker_Symbol()
Dim ws As Worksheet
Dim Ticker As String
Dim yearly_change As Double
Dim openprice As Double
Dim closingprice As Double
Dim percent As Double
Dim volume As Double
Dim lastrow As Double

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

volume = 0

openprice = Cells(2, 3).Value

summary_table_row = 2

For i = 2 To lastrow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Ticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
       
        closingprice = Cells(i, 6).Value
        
        yearly_change = closingprice - openprice
        
        If openprice = 0 Then
                percent = 0
            Else
            percent = yearly_change / openprice
            End If
        
        ws.Cells(summary_table_row, 9).Value = Ticker
        ws.Cells(summary_table_row, 10).Value = yearly_change
        ws.Cells(summary_table_row, 11).Value = percent
        ws.Cells(summary_table_row, 12).Value = volume
        
        
        openprice = Cells(i + 1, 3).Value
      
        volume = 0
        
    If yearly_change > 0 Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    
    ElseIf yearly_change < 0 Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
    
    End If
     summary_table_row = summary_table_row + 1
Else
    
    volume = volume + Cells(i, 7).Value

    End If
    
Next i

Next ws



End Sub
