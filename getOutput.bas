Function Get_Unique_Values()
    Dim row As Long
    
    row = Cells(Rows.Count, "A").End(xlUp).row
    ActiveSheet.Range("A2:A" & row).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("L2"), Unique:=True
End Function

Sub getOutput()
   
   On Error Resume Next
   
   Dim minDate As Date
   Dim maxDate As Date
   Dim minOpenValue As Double
   Dim maxCloseValue As Double
   
   Dim lDate As Date
   
   Dim outYrChange As Double
   Dim outPcentChange As Double
   Dim outVolume As Long
   
   Dim jMax As Integer
   Dim iMax As Long
   Dim s As String
   Dim ws As Worksheet
   Dim year As String
   Dim bOutVolume As Boolean
   
   For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Set ws = ActiveSheet
      ws.Range("L1").Value = "Ticker"
      ws.Range("M1").Value = "Yearly Change"
      ws.Range("N1").Value = "Percent Change"
      ws.Range("O1").Value = "Total Stock volume"
   
      year = ws.Name
       ' get unique tickers
      Get_Unique_Values
    
      jMax = ws.Cells(Rows.Count, "L").End(xlUp).row  ' number of unique ticker
      iMax = ws.Cells(Rows.Count, "A").End(xlUp).row  ' all rows in worksheet
      
      'Calculate other values for each ticker
      For j = 2 To jMax
         outVolume = 0
         bOutVolume = True
         minDate = CDate("12/31/" + year)
         maxDate = CDate("01/01/" + year)
    
         For i = 2 To iMax
            If ws.Cells(i, 1).Value = ws.Cells(j, 12) Then
               If bOutVolume Then
                  outVolume = outVolume + ws.Cells(i, 7).Value
               
                  If Err.Number = 6 Then
                     bOutVolume = False
                     Err.Clear
                  End If
               End If
               'convert column 2 to date and check if its lower than mindate
               'if lower than minDate copy opening value
               'if greater than maxDate copy closing value
         
               s = ws.Cells(i, 2).Value
               lDate = DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 5, 2)), CInt(Right(s, 2)))
         
               If lDate < minDate Then
                  minDate = lDate
                  minOpenValue = ws.Cells(i, 3).Value
            
               ElseIf lDate > maxDate Then
                  maxDate = lDate
                  maxCloseValue = ws.Cells(i, 6).Value
               End If
            End If
         Next i
    
         outYrChange = (maxCloseValue - minOpenValue) ' Yearly change
         ws.Cells(j, 13).Value = outYrChange
    
         'conditional formatting for Yearly Change column
         If outYrChange > 0 Then
            ws.Cells(j, 13).Interior.ColorIndex = 4
         ElseIf outYrChange < 0 Then
            ws.Cells(j, 13).Interior.ColorIndex = 3
         End If
    
         outPcentChange = outYrChange / minOpenValue
         ws.Cells(j, 14).Value = FormatPercent(outPcentChange, 2)
    
         'conditional formatting for Percent Change column
         If outPcentChange > 0 Then
            ws.Cells(j, 14).Interior.ColorIndex = 4
         ElseIf outPcentChange < 0 Then
            ws.Cells(j, 14).Interior.ColorIndex = 3
         End If
    
         If bOutVolume Then
            ws.Cells(j, 15).Value = outVolume
         End If
      Next j
   
      ' get number of rows in new table
      n = ws.Range("N1", Range("N1").End(xlDown)).Rows.Count
      
      ws.Range("R3").Value = "Greatest % Increase"
      ws.Range("R4").Value = "Greatest % Decrease"
      ws.Range("R5").Value = "Greatest Total Volume"
      ws.Range("S2").Value = "Ticker"
      ws.Range("T2").Value = "Value"
   
      Dim highestYrIncrease As Double
      Dim highestYrDecrease As Double
      Dim highestVolume As Long
   
      highestYrIncrease = 0
      highestYrDecrease = 0
      highestVolume = 0
   
      For i = 2 To n
         If Cells(i, 14).Value > highestYrIncrease Then  ' column N Percent change
            Range("S3").Value = Cells(i, 12).Value       ' column L ticker
            Range("T3").Value = FormatPercent(Cells(i, 14).Value, 1)       ' column N Percent change
            highestYrIncrease = Cells(i, 14).Value       ' next time compare with highest value found till now
         
         ElseIf Cells(i, 14).Value < highestYrDecrease Then  ' column N Percent change
            Range("S4").Value = Cells(i, 12).Value           ' column L ticker
            Range("T4").Value = FormatPercent(Cells(i, 14).Value, 1)           ' column N Percent change
            highestYrDecrease = Cells(i, 14).Value           ' next time compare with lowest value found till now
         End If
      
         If Cells(i, 15).Value > highestVolume Then  ' column O total stock volume
            Range("S5").Value = Cells(i, 12).Value   ' column L ticker
            Range("T5").Value = Cells(i, 15).Value   ' column O total stock volume
            highestVolume = Cells(i, 15).Value       ' next time compare with lowest value found till now
         End If
      Next i
   
   Next ws
End Sub

