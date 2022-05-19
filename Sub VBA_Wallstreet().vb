Sub VBA_Wallstreet()

For Each ws In Worksheets
'Column Headers
ws.Range("I1") = "<Ticker>"
ws.Range("J1") = "<Yearly Change>"
ws.Range("K1") = "<Percent Change>"
ws.Range("L1") = "<Total Stock Volume>"

'Variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Volume As Double
Dim StockOpen As Double
Dim StockClose As Double

Sum_Table_Row = 2
'determine last row
LRow = Cells(Rows.Count, 1).End(xlUp).Row
Volume = 0

' For Loop
For x = 2 To LRow
        
    'Finding Ticker and Volume
    If ws.Cells(x + 1, 1) <> ws.Cells(x, 1) Then
        Ticker = ws.Cells(x, 1).Value
        Volume = Volume + ws.Cells(x, 7).Value
        
        ws.Range("I" & Sum_Table_Row).Value = Ticker
        ws.Range("L" & Sum_Table_Row).Value = Volume
        
        Volume = 0
        StockClose = ws.Cells(x, 6)
        
        'Stock change
        If StockOpen = 0 Then
        YearlyChange = 0
        PercentChange = 0
        
        Else:
        YearlyChange = StockClose - StockOpen
        PercentChange = (StockClose - StockOpen) / StockOpen
        
        End If
        
        'Percentage Change
        ws.Range("J" & Sum_Table_Row).Value = YearlyChange
        ws.Range("K" & Sum_Table_Row).Value = PercentChange
        ws.Range("K" & Sum_Table_Row).Style = "Percent"
        ws.Range("K" & Sum_Table_Row).NumberFormat = "0.00%"
        
        Sum_Table_Row = Sum_Table_Row + 1
        
    'Identifying Stock Opening Value
    ElseIf ws.Cells(x - 1, 1).Value <> ws.Cells(x, 1) Then
        StockOpen = ws.Cells(x, 3)
    
    Else: Volume = Volume + ws.Cells(x, 7).Value
    End If
    Next x

'Conditional formating for yearly change.
For Z = 2 To LRow
    If ws.Range("J" & Z).Value > 0 Then
    ws.Range("J" & Z).Interior.ColorIndex = 4
    
     ElseIf ws.Range("J" & Z).Value < 0 Then
        ws.Range("J" & Z).Interior.ColorIndex = 3
        
    End If
    Next Z
   
'Bonus content
'Setting cell range to populate with answers
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Bonus Variables
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

'Ensuring each starts at zero
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

'Find greatest increase
For a = 2 To LRow

    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

    Next a

'Find greatest decrease
For b = 2 To LRow
    
    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
   Next b

'Find greatest volume
For c = 2 To LRow
    
    If ws.Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
    Next c

Next ws
End Sub


  
        



