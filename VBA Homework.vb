Sub ticker_review()

'loop through each worksheet

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

'define last row

lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set Headers

Cells(1, 9).Value = "Ticker"
Cells(1, 12).Value = "Total Stock Volume"


  ' Define variables
  
  Dim ticker As String
  Dim ticker_total As Double
  ticker_total = 0

  ' Summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For i = 2 To lr

    ' Conditional Function
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value
      ticker_total = ticker_total + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = ticker
      Range("L" & Summary_Table_Row).Value = ticker_total

      Summary_Table_Row = Summary_Table_Row + 1
      ticker_total = 0
    Else
     ticker_total = ticker_total + Cells(i, 7).Value

    End If

  Next i
  
  
  Next ws
  
End Sub
