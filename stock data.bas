Attribute VB_Name = "Module1"
Sub Stockdata():
    Dim ws As Worksheet



For Each ws In ThisWorkbook.Worksheets
ws.Activate


    Dim i As Double
    Dim ticker As String
    Dim summaryrow As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim total As Double
    Dim max_inc As Double
    Dim max_dec As Double
    Dim max_vol As Double
    
  
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    summaryrow = 2
    price_row = 2
  

  For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
  If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
  

    ticker = Cells(i, 1).Value
    Cells(summaryrow, 9).Value = ticker
  

  
    openprice = Cells(price_row, 3).Value
    
    closeprice = Cells(i, 6).Value
    
    yearlychange = closeprice - openprice
    
    Cells(summaryrow, 10).Value = yearlychange
  
  

   Cells(summaryrow, 11).Value = Round((yearlychange / openprice), 4)
   
   percentchange = Cells(summaryrow, 11).Value
   

 
     total = total + Cells(i, 7).Value
     Cells(summaryrow, 12).Value = total


    summaryrow = summaryrow + 1
    price_row = i + 1
    total = 0
     
  Else
  
    total = total + Cells(i, 7).Value
 
  End If
  
  If Cells(i, 10).Value < 0 Then
  Cells(i, 10).Interior.ColorIndex = 3
  Else
  Cells(i, 10).Interior.ColorIndex = 4
  End If
  
  If Cells(i, 11).Value = Application.Max(Range("K2:K3001")) Then
  Range("Q2").Value = Cells(i, 11).Value
  Range("P2").Value = Cells(i, 9).Value
  ElseIf Cells(i, 11).Value = Application.Min(Range("K2:K3001")) Then
  Range("Q3").Value = Cells(i, 11).Value
  Range("P3").Value = Cells(i, 9).Value
  End If
  
  If Cells(i, 12).Value = Application.Max(Range("L2:L3001")) Then
  Range("Q4").Value = Cells(i, 12).Value
  Range("P4").Value = Cells(i, 9).Value
  End If
  
  Next i
  Next ws
  
 
End Sub
