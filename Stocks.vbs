Sub Stocks()
    Dim volume As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim row_count As Long
    Dim percent_change As Double
    Dim ws As Worksheet
    For Each ws In Worksheets
       j = 2
       start = 2
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       row_count = Cells(rows.Count, "A").End(xlUp).Row
       For i = 2 To row_count
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               volume = volume + ws.Cells(i, 7).Value
               If volume = 0 Then
                   ws.Range("I" & j).Value = Cells(i, 1).Value
                   ws.Range("J" & j).Value = 0
                   ws.Range("K" & j).Value = "%" & 0
                   ws.Range("L" & j).Value = 0
               Else
                   If ws.Cells(start, 3) = 0 Then
                       For find_value = start To i
                           If ws.Cells(find_value, 3).Value <> 0 Then
                               start = find_value
                               Exit For
                           End If
                       Next find_value
                   End If
                   change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                   percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                   start = i + 1
                   ws.Range("I" & j).Value = ws.Cells(i, 1).Value
                   ws.Range("J" & j).Value = Round(change, 2)
                   ws.Range("K" & j).Value = "%" & percentChange
                   ws.Range("L" & j).Value = volume
                   Select Case change
                       Case Is > 0
                           ws.Range("J" & j).Interior.ColorIndex = 4
                       Case Is < 0
                           ws.Range("J" & j).Interior.ColorIndex = 3
                       Case Else
                           ws.Range("J" & j).Interior.ColorIndex = 0
                   End Select
               End If
               volume = 0
               change = 0
               j = j + 1
           Else
               volume = volume + ws.Cells(i, 7).Value
           End If
       Next i
   Next ws
End Sub
