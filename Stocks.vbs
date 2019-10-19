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
       ' Set title row
        ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       ' get the row number of the last row with data
       row_count = Cells(rows.Count, "A").End(xlUp).Row
       For i = 2 To row_count
           ' If ticker changes then print results
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               ' Stores results in variables
               volume = volume + ws.Cells(i, 7).Value
               ' Handle zero total volume
               If volume = 0 Then
                   ' print the results
                   ws.Range("I" & j).Value = Cells(i, 1).Value
                   ws.Range("J" & j).Value = 0
                   ws.Range("K" & j).Value = "%" & 0
                   ws.Range("L" & j).Value = 0
               Else
                   ' Find First non zero starting value
                   If ws.Cells(start, 3) = 0 Then
                       For find_value = start To i
                           If ws.Cells(find_value, 3).Value <> 0 Then
                               start = find_value
                               Exit For
                           End If
                       Next find_value
                   End If
                   ' Calculate Change
                   change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                   percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                   ' start of the next stock ticker
                   start = i + 1
                   ' print the results to a separate worksheet
                   ws.Range("I" & j).Value = ws.Cells(i, 1).Value
                   ws.Range("J" & j).Value = Round(change, 2)
                   ws.Range("K" & j).Value = "%" & percentChange
                   ws.Range("L" & j).Value = volume
                   ' colors positives green and negatives red
                   Select Case change
                       Case Is > 0
                           ws.Range("J" & j).Interior.ColorIndex = 4
                       Case Is < 0
                           ws.Range("J" & j).Interior.ColorIndex = 3
                       Case Else
                           ws.Range("J" & j).Interior.ColorIndex = 0
                   End Select
               End If
               ' reset variables for new stock ticker
               volume = 0
               change = 0
               j = j + 1
           ' If ticker is still the same add results
           Else
               volume = volume + ws.Cells(i, 7).Value
           End If
       Next i
   Next ws
End Sub
