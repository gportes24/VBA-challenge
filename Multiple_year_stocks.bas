Attribute VB_Name = "Module1"
Sub Stockticker()
    Dim ticker As String
    Dim total_vol As Double
    Dim Summary_Table_Row As Integer
    Range("H1").Value = "Ticker"
    Range("Q1").Value = "ticker"
    Range("R1").Value = "Value"
    Range("I1").Value = "Yearly Change"
    Range("J1").Value = "Percentage Change"
    Range("K1").Value = "Total Stock Volume"
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Summary_Table_Row = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_Row = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    total_vol = 0
    stock_open = Cells(2, 3).Value
    Range("J2:J" & lastRow).NumberFormat = "0.00%"

    For Row = 2 To lastRow

    If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
      ticker = Cells(Row, 1).Value
      total_vol = total_vol + Cells(Row, 7).Value
      stock_close = Cells(Row, 6).Value
      stock_yr_change = stock_close - stock_open
      stock_per_change = (stock_yr_change / stock_open)
      Range("H" & Summary_Table_Row).Value = ticker
      Range("I" & Summary_Table_Row).Value = stock_yr_change
      Range("J" & Summary_Table_Row).Value = stock_per_change
      Range("K" & Summary_Table_Row).Value = total_vol
      Summary_Table_Row = Summary_Table_Row + 1

      total_vol = 0

    Else
        total_vol = total_vol + Cells(Row, 7).Value

    End If

    Next Row
 'stackoverflow help
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For row2 = 1 To c
    Set color_cell = rg(row2)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next row2
End Sub

