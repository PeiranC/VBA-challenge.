Sub stock_data()

 Dim ws As Long
 Dim ws_Count As Long
 
 
 'loop through all worksheet
 
 ws_Count = ActiveWorkbook.Worksheets.Count
 
For ws = 1 To ws_Count


'name the summary tabel title



Worksheets(ws).Cells(1, "I") = "Ticker"

Worksheets(ws).Cells(1, "J") = "Yearly Change"

Worksheets(ws).Cells(1, "K") = "Percentage Change"

Worksheets(ws).Cells(1, "L") = "Total Stock Volume"

Worksheets(ws).Columns("A:Q").AutoFit



lastRow = Worksheets(ws).Cells(Rows.Count, "A").End(xlUp).Row


Total = 0
First_Open_Price = 2
Summary_Poiner = 2

For i = 2 To lastRow


If Worksheets(ws).Cells(i, "A").Value <> Worksheets(ws).Cells(i + 1, "A").Value Then

'calucate A to G

Total = Total + Worksheets(ws).Cells(i, "G").Value

Open_Price = Worksheets(ws).Cells(First_Open_Price, "C").Value

Close_Price = Worksheets(ws).Cells(i, "F").Value

Yearly_change = Close_Price - Open_Price

Percentage_Change = Yearly_change / Open_Price * 100

'Get result

Worksheets(ws).Cells(Summary_Poiner, "I").Value = Worksheets(ws).Cells(i, "A").Value
Worksheets(ws).Cells(Summary_Poiner, "J").Value = Yearly_change
Worksheets(ws).Cells(Summary_Poiner, "K").Value = "%" & Percentage_Change
Worksheets(ws).Cells(Summary_Poiner, "L").Value = Total

'put conditions

If Yearly_change > 0 Then

Worksheets(ws).Cells(Summary_Poiner, "J").Interior.ColorIndex = 4

Else
 
Worksheets(ws).Cells(Summary_Poiner, "J").Interior.ColorIndex = 3
 
End If


Total = 0

First_Open_Price = i + 1

Summary_Poiner = Summary_Poiner + 1

Else

Total = Total + Worksheets(ws).Cells(i, "G").Value

End If

Next i

Next ws

End Sub

Sub stock_change()

 Dim ws As Long
 Dim ws_Count As Long
 
 ws_Count = ActiveWorkbook.Worksheets.Count
 
For ws = 1 To ws_Count

Worksheets(ws).Cells(1, "P").Value = "Ticker"
Worksheets(ws).Cells(1, "Q").Value = "Value"
Worksheets(ws).Cells(2, "O").Value = "Greatest % Increase"
Worksheets(ws).Cells(3, "O").Value = "Greatest % Decrease"
Worksheets(ws).Cells(4, "O").Value = "Greatest Total Volume"
Worksheets(ws).Columns("O").AutoFit

LastRow2 = Worksheets(ws).Cells(Rows.Count, "I").End(xlUp).Row

For j = 2 To LastRow2

Greatest_Increase = Application.WorksheetFunction.Max(Worksheets(ws).Range("K:K"))
Greatest_Decrease = Application.WorksheetFunction.Min(Worksheets(ws).Range("K:K"))
Top_Vol = Application.WorksheetFunction.Max(Worksheets(ws).Range("L:L"))



If Worksheets(ws).Cells(j, "K").Value = Greatest_Increase Then

Worksheets(ws).Cells(2, "P").Value = Worksheets(ws).Cells(j, "I").Value
Worksheets(ws).Cells(2, "Q").Value = "%" & Worksheets(ws).Cells(j, "K").Value

ElseIf Worksheets(ws).Cells(j, "K").Value = Greatest_Decrease Then

Worksheets(ws).Cells(3, "P").Value = Worksheets(ws).Cells(j, "I").Value

Worksheets(ws).Cells(3, "Q").Value = "%" & Worksheets(ws).Cells(j, "K").Value


ElseIf Worksheets(ws).Cells(j, "L").Value = Top_Vol Then

Worksheets(ws).Cells(4, "P").Value = Worksheets(ws).Cells(j, "I").Value
Worksheets(ws).Cells(4, "Q").Value = Worksheets(ws).Cells(j, "L").Value



End If
            
Next j
           
           
Next ws

    

End Sub
