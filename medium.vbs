Sub medium()

'Define Variable
  
Dim Closing_Price As Double

Dim Open_Price As Double
   
Dim Yearly_Price As Double

Dim Percentage_Change As Double

Dim Total_Stock_Volume As Double

'Inputing name title for Row 1




'Define Last Row

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"

ws.Range("J1").Value = "Yearly Change"

ws.Range("k1").Value = "Percent Change"

ws.Range("L1").Value = "Total Stock Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Lastrow_2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

Table_Row = 2

'Running for Loop

Open_Price = ws.Cells(2, 3).Value



    For i = 2 To LastRow

      


        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
        ws.Range("I" & Table_Row) = ws.Cells(i, 1).Value

   Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value


        Closing_Price = ws.Cells(i, 6).Value

        Yearly_Change = Closing_Price - Open_Price

                If Open_Price = 0 Then
                    
                    Open_Price = ws.Cells(i + 2, 3).Value

                    Percentage_Change = (Closing_Price - Open_Price) / Open_Price

                Else

                    Percentage_Change = (Closing_Price - Open_Price) / Open_Price

                End If



        
        ws.Range("J" & Table_Row).Value = Yearly_Change

        ws.Range("K" & Table_Row).Value = FormatPercent(Percentage_Change)

        ws.Range("L" & Table_Row).Value = Total_Stock_Volume

        Open_Price = ws.Cells(i + 1, 3).Value
        
        
        Table_Row = Table_Row + 1

          Total_Stock_Volume = 0
        Else

                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If

    Next i
    


    For j = 2 To Lastrow_2

        If ws.Cells(j, 10).Value >= 0 Then

        ws.Cells(j, 10).Interior.ColorIndex = 4

        Else

        ws.Cells(j, 10).Interior.ColorIndex = 3

        End If

    Next j
    

Next ws

End Sub

