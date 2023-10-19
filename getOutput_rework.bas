Attribute VB_Name = "Module1"
Sub getOutput()
' Set initial values
On Error Resume Next

   Dim j As Integer
   Dim total As Long
   Dim change As Double
   Dim start As Long
   Dim rowCount As Long
   Dim Rng As Range
   Dim Mx As Double
   Dim Rw As Integer
   
   
   
   For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Set ws = ActiveSheet
      
      ws.Range("I1").Value = "<Ticker>"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Stock Volume"
      
      j = 0
      total = 0
      change = 0
      start = 2

      ' get the row number of the last row with data
      rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row

      For i = 2 To rowCount

         ' If ticker changes then print results
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + ws.Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

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
                percentChange = change / ws.Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            Days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value
            If Err.Number = 6 Then
               Err.Clear
            End If
        End If

    Next i
    
    ' Get highest increase, decrease and stock volume
    
    rowCount = ws.Cells(Rows.Count, "K").End(xlUp).row
    Set Rng = Range("K2:K" & rowCount)
    Mx = WorksheetFunction.Max(Rng)
    Rw = WorksheetFunction.Match(Mx, Rng, 0) + Rng.row - 1
    
    ws.Cells(2, 16).Value = ws.Cells(Rw, 9).Value
    ws.Cells(2, 17).Value = Mx
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ' Get highest increase, decrease and stock volume
    
    rowCount = ws.Cells(Rows.Count, "K").End(xlUp).row
    Set Rng = Range("K2:K" & rowCount)
    Mx = WorksheetFunction.Min(Rng)
    Rw = WorksheetFunction.Match(Mx, Rng, 0) + Rng.row - 1
    
    ws.Cells(3, 16).Value = ws.Cells(Rw, 9).Value
    ws.Cells(3, 17).Value = Mx
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Get highest increase, decrease and stock volume
    
    rowCount = ws.Cells(Rows.Count, "L").End(xlUp).row
    Set Rng = Range("L2:L" & rowCount)
    Mx = WorksheetFunction.Max(Rng)
    Rw = WorksheetFunction.Match(Mx, Rng, 0) + Rng.row - 1
    
    ws.Cells(4, 16).Value = ws.Cells(Rw, 9).Value
    ws.Cells(4, 17).Value = Mx
    
   Next ws
End Sub


