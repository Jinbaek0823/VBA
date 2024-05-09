# VBA_Homework
Sub QuarterlyChangeForAllSheets()

    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim J As Integer
    Dim start As Long
    Dim Rowcount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    For Each ws In ThisWorkbook.Sheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quartely Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greastest Total Volume"
        
 
        total = 0
        J = 0
        change = 0
        start = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        Rowcount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To Rowcount

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                percentChange = (change / ws.Cells(start, 3).Value) * 100
                
     
                ws.Range("I" & 2 + J).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + J).Value = change
                ws.Range("K" & 2 + J).Value = percentChange
                ws.Range("L" & 2 + J).Value = total

                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerIncrease = ws.Cells(i, 1).Value
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerDecrease = ws.Cells(i, 1).Value
                End If
                
                If total > greatestVolume Then
                    greatestVolume = total
                    tickerVolume = ws.Cells(i, 1).Value
                End If
                
                total = 0
                J = J + 1
                start = i + 1
            Else
  
                total = total + ws.Cells(i, 7).Value
            End If
            
        Next i
        

        ws.Range("P2").Value = tickerIncrease
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("P3").Value = tickerDecrease
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("P4").Value = tickerVolume
        ws.Range("Q4").Value = greatestVolume
    Next ws
    
End Sub


