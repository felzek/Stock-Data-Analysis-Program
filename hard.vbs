Sub medium()

'Define Variable
  
Dim Closing_Price As Double

Dim Open_Price As Double
   
Dim Yearly_Price As Double

Dim Percentage_Change As Double

Dim Total_Stock_Volume As Double

'Inputing name title for Row 1

ws.Range("I1").Value = "Ticker"

ws.Range("J1").Value = "Yearly Change"

ws.Range("k1").Value = "Percent Change"

ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"

ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"

ws.Range("O3").Value = "Greatest % Decrease"

ws.Range("O4").Value = "Greatest Total Volume"

'Define Last Row

For Each ws In Worksheets

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

        volume = ws.Cells(i, 7).Value
      

   

               

                If Open_Price = 0 Then

                 Open_Price = ws.Cells(i + 2, 3).Value

                    Percentage_Change = (Closing_Price - Open_Price) / Open_Price
                Else

                    Percentage_Change = (Closing_Price - Open_Price) / Open_Price

                End If

     Yearly_Change = Closing_Price - Open_Price

        
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
    


    Dim max_year_change As Integer

    max_increase = ws.Cells(2, 11).Value

    max_decrease = ws.Cells(2, 11).Value

    max_vol = ws.Cells(2, 12).Value
    
    greatest_volume = ws.Cells(2, 12).Value

    Table_Row_2 = 0

    For i = 2 To Lastrow_2

        If max_increase < ws.Cells(i, 11).Value Then

            max_increase = ws.Cells(i, 11).Value

       
          

        End If
 
           
        If max_decrease > ws.Cells(i, 11).Value Then

              max_decrease = ws.Cells(i, 11).Value

         

            
        End If
        
        If greatest_volume < ws.Cells(i, 12).Value Then

            greatest_volume = ws.Cells(i, 12).Value

        End If


 
            ws.Range("Q2").Value = FormatPercent(max_increase)
            ws.Range("Q3").Value = FormatPercent(max_decrease)
            ws.Range("Q4").Value = greatest_volume

               If ws.Cells(i, 11) = ws.Range("Q2") Then
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            End If
 If ws.Cells(i, 11) = ws.Range("Q3") Then
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 12) = ws.Range("Q4") Then
                ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
    Next i

    max_increase = 0
    max_decrease = 0

        

Next ws

End Sub
