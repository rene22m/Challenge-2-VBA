Sub LoopThroughEverySheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Challenge2Stocks
    Next
    Application.ScreenUpdating = True
End Sub
Sub Challenge2Stocks()

  Dim Ticker, year_date As String
  Dim TotalVolume, init As Variant
  Dim Summary_Table_Row As Integer
  Dim n As Long
  
  TotalVolume = 0
  Summary_Table_Row = 2
  n = Worksheets("2018").UsedRange.Rows.Count
  init = 2
  For i = 2 To n
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = TotalVolume
        TotalVolume = 0
        close_price = Cells(i, 6).Value
        open_price = Cells(init, 3).Value
        Change = close_price - open_price
        Cells(Summary_Table_Row, 10).Value = Change
            If Change >= 0 Then
            Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
            Else
            Cells(Summary_Table_Row, 10).Interior.Color = vbRed
            End If
        PercentChange = ((close_price - open_price) / open_price) * 100
        Cells(Summary_Table_Row, 11).Value = PercentChange
        Summary_Table_Row = Summary_Table_Row + 1
        init = i + 1
        
    Else
    TotalVolume = TotalVolume + Cells(i, 7).Value

    End If
Next i

Columns("J:J").Select
    Selection.Style = "Currency"
    
    Range("Q2").FormulaR1C1 = "=MAX(C[-6])"
    Range("Q3").FormulaR1C1 = "=MIN(C[-6])"
    Range("Q4").FormulaR1C1 = "=MAX(C[-5])"

    Range("P2").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
    Range("P3").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-5],0))"
    Range("P4").FormulaR1C1 = "=INDEX(C[-7],MATCH(RC[1],C[-4],0))"

    Range("I1").FormulaR1C1 = "Ticker"
    Columns("I:I").EntireColumn.AutoFit
    Range("J1").FormulaR1C1 = "Yearly Change"
    Columns("J:J").EntireColumn.AutoFit
    Range("K1").FormulaR1C1 = "Percent Change"
    Columns("K:K").EntireColumn.AutoFit
    Range("L1").FormulaR1C1 = "Total Stock Volume"
    Columns("L:L").EntireColumn.AutoFit
    Range("O2").FormulaR1C1 = "Greatest % Increase"
    Range("O3").FormulaR1C1 = "Greatest % Decrease"
    Range("O4").FormulaR1C1 = "Greatest Total Volume"
    Columns("O:O").EntireColumn.AutoFit
    Range("P1").FormulaR1C1 = "Ticker"
    Range("Q1").FormulaR1C1 = "Value"
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    
End Sub