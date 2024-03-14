'code taken from class material. Worked with Astrid and Paola on assignment

Sub tickervolume()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim lastrow As Long
        Dim tickersum As String
        Dim summary_table_row As Long
        Dim volume As Double
            summary_table_row = 2
            volume = 0
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                    tickersum = ws.Cells(i, 1).value
                    volume = volume + ws.Cells(i, 7).value
                    'Cells(summary_table_Row, 9).Value = tickersum
                    ws.Range("I" & summary_table_row).value = tickersum
                    ws.Range("L" & summary_table_row).value = volume
                    summary_table_row = summary_table_row + 1
                    volume = 0
            Else
                    volume = volume + ws.Cells(i, 7).value
                        
            End If
        Next i
    Next ws
End Sub

Sub opencloseprice()
    Dim ws As Worksheet
    For Each ws In Worksheets
        MsgBox (ws.Name)
        
        Dim lastrow As Long
        Dim openprice As Double
        Dim closeprice As Double
        Dim summary_table_row As Integer
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summary_table_row = 2
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                closeprice = ws.Cells(i, 6).value
                ws.Range("R" & summary_table_row).value = closeprice
                summary_table_row = summary_table_row + 1
            End If
            If ws.Cells(i, 1).value <> ws.Cells(i - 1, 1).value Then
                openprice = ws.Cells(i, 3).value
                ws.Range("Q" & summary_table_row).value = openprice
        
            End If
                
        Next i
    Next ws
End Sub

Sub yearlychange()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearchange As Double
        Dim lastrow As Long
        Dim percentchange As Double
        Dim summary_table_row As Integer
        
        summary_table_row = 2
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow
            openprice = ws.Cells(i, 17).value
            closeprice = ws.Cells(i, 18).value
            yearchange = closeprice - openprice
            ws.Range("J" & summary_table_row).value = yearchange
        
                 If openprice <> 0 Then
                    percentchange = (yearchange / openprice)
                Else
                    percentchange = 0
                End If

        
            ws.Range("K" & summary_table_row).value = percentchange
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            summary_table_row = summary_table_row + 1
            
            If ws.Range("J" & i).value < 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 3
            Else
                ws.Range("J" & i).Interior.ColorIndex = 4
            End If
            
            If ws.Range("K" & i).value < 0 Then
                ws.Range("K" & i).Interior.ColorIndex = 3
            Else
                ws.Range("K" & i).Interior.ColorIndex = 4
            End If
        Next i
    Next ws
End Sub
' code created using online references and help from learning assistant
Sub percentsum()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim ticker As String
        Dim greatpercentvalue As Double
        Dim leastpercentvalue As Double
        Dim greatestvolume As Double
        Dim lastrowpercent As Long
        Dim lastrowvolume As Long
        
        lastrowvolume = ws.Cells(Rows.Count, 12).End(xlUp).Row
        lastrowpercent = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        greatpercentvalue = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 23).value = greatpercentvalue
        ws.Cells(2, 23).NumberFormat = "0.00%"
        
        
        leastpercentvalue = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 23).value = leastpercentvalue
        ws.Cells(3, 23).NumberFormat = "0.00%"
        
        greatestvolume = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 23).value = greatestvolume
        
            For i = 2 To lastrowpercent
                If ws.Cells(i, 11).value = greatpercentvalue Then
                    ws.Cells(2, 22).value = ws.Range("I" & i)
                    ws.Cells(2, 21).value = "greatest % increase"
                ElseIf ws.Cells(i, 11).value = leastpercentvalue Then
                    ws.Cells(3, 22).value = ws.Range("I" & i)
                    ws.Cells(3, 21).value = "greatest % decrease"
                End If
            Next i
                      
               For i = 2 To lastrowvolume
                    If ws.Cells(i, 12).value = greatestvolume Then
                    ws.Cells(4, 22).value = ws.Range("I" & i)
                    ws.Cells(4, 21).value = "greatest total volume"
                    End If
                Next i
    Next ws
End Sub


