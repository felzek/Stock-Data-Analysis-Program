Sub ticker()

'declare variables

Dim ticker As String
Dim TotalVol As Double
Dim ws As Worksheet

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


For Each ws In ActiveWorkbook.Worksheets

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow


        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ticker = ws.Cells(i, 1).Value

            TotalVol = TotalVol + ws.Cells(i, 7).Value

            ws.Range("I" & Summary_Table_Row).Value = ticker

            ws.Range("J" & Summary_Table_Row).Value = TotalVol

            Summary_Table_Row = Summary_Table_Row + 1

            TotalVol = 0

        Else

            TotalVol = TotalVol + ws.Cells(i, 7).Value

        End If

    Next i


Next ws

End Sub
