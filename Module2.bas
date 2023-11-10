Attribute VB_Name = "Module2"
Sub Ticker()

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Change"
ws.Range("O1").Value = "Greatest % increase"
ws.Range("O2").Value = "Greatest % decrease"
ws.Range("O3").Value = "Greatest total volume"


Dim Lastrow As Long
    Lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
   
  Dim Ticker As String
  Dim ticker_count As Double
  Dim Percent_Changee As Double
  Percent_Change = 0
  Dim Yearly_Change As Double
  Yearly_Change = 0
  Dim Volume_Total As Double
  Volume_Total = 0

  Dim YearlyChange_Row As Integer
  YearlyChange_Row = 2
  Dim Percent_Change_Row As Integer
  Percent_Change_Row = 2
  Dim Volume_Total_Row As Integer
  Volume_Total_Row = 2
  Open_Ticker = ws.Cells(2, 3).Value
  
For i = 2 To Lastrow

     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value
      ws.Range("I" & YearlyChange_Row).Value = Ticker
      
      Yearly_Change = ws.Cells(i, 6) - ws.Cells(i - ticker_count, 3).Value
      ws.Range("J" & YearlyChange_Row).Value = Yearly_Change
      
      Volume_Total = Volume_Total + Cells(i, 7).Value
      ws.Range("L" & YearlyChange_Row).Value = Volume_Total
      Percent_Change = Yearly_Change / ws.Cells(i - ticker_count, 3).Value
      ws.Range("K" & YearlyChange_Row).Value = Percent_Change
      ws.Range("K" & YearlyChange_Row).NumberFormat = "0.00%"
      
      YearlyChange_Row = YearlyChange_Row + 1
      Yearly_Change = 0
      Volume_Total = 0
      ticker_count = 0

    Else

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      ticker_count = ticker_count + 1

    End If
     
Next i
   
   Dim sum_ticker As String
   Dim max_res As Double
   max_res = 0
   Dim min_res As Double
   min_res = 0
   Dim gre_vol As Double
   gre_vol = 0
   Dim total_lastrow As Long
   total_lastrow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
   
For c = 2 To total_lastrow
      
  If ws.Cells(c, 11).Value >= max_res Then
   sum_ticker = ws.Cells(c, 9).Value
   max_res = ws.Cells(c, 11).Value
   ws.Range("P1").Value = sum_ticker
   ws.Range("Q1").Value = max_res
   
   ElseIf ws.Cells(c, 12).Value >= gre_vol Then
   sum_ticker = ws.Cells(c, 9).Value
   gre_vol = ws.Cells(c, 12).Value
   ws.Range("P3").Value = sum_ticker
   ws.Range("Q3").Value = gre_vol
   
   ElseIf ws.Cells(c, 11).Value <= min_res Then
   sum_ticker = ws.Cells(c, 9).Value
   min_res = ws.Cells(c, 11).Value
   ws.Range("P2").Value = sum_ticker
   ws.Range("Q2").Value = ws.Cells(c, 11).Value
    
   Else
   End If
   
   If ws.Cells(c, 11).Value >= 0 Then
    ws.Cells(c, 11).Interior.Color = RGB(0, 255, 0)
    Else
    ws.Cells(c, 11).Interior.Color = RGB(255, 0, 0)
    End If
   
   If ws.Cells(c, 10).Value >= 0 Then
    ws.Cells(c, 10).Interior.Color = RGB(0, 255, 0)
    Else
    ws.Cells(c, 10).Interior.Color = RGB(255, 0, 0)
    End If
    ws.Range("Q1", "Q2").NumberFormat = "0.00%"
Next c
    
Next ws

End Sub

