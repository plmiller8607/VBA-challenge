Sub StockAnalysis()
    'Set dimensions
    Dim total As Double
    Dim rowindex As Long
    Dim change As Double
    Dim columnindex As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim dailyChange As Single
    Dim averageChange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        columnindex = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        'Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Total Volume"
        
        'Get the row number of the last row with data (because the number is dynamic)
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For rowindex = 2 To rowCount
            
            'If ticker changes then print the results
            If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
            
                'Store results in a variable
                total = total + ws.Cells(rowindex, 7).Value
                
                If total = 0 Then
                
                    'Print the results
                    ws.Range("i" & 2 + columnindex).Value = Cells(rowindex, 1).Value
                    ws.Range("j" & 2 + columnindex).Value = 0
                    ws.Range("k" & 2 + columnindex).Value = "%" & 0
                    ws.Range("l" & 2 + columnindex).Value = 0
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To rowindex
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                                                        
                        Next find_value
                                        
                    End If
                    change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)
                    
                    start = rowindex + 1
                    ws.Range("i" & 2 + columnindex) = ws.Cells(rowindex, 1).Value
                    ws.Range("j" & 2 + columnindex) = change
                    ws.Range("j" & 2 + columnindex).NumberFormat = "0.00"
                    ws.Range("k" & 2 + columnindex).Value = percentChange
                    ws.Range("k" & 2 + columnindex).NumberFormat = "0.00%"
                    ws.Range("l" & 2 + columnindex).Value = total
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("j" & 2 + columnindex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("j" & 2 + columnindex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("j" & 2 + columnindex).Interior.ColorIndex = 0
                    End Select
                     
                    Select Case change
                        Case Is > 0
                            ws.Range("k" & 2 + columnindex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("k" & 2 + columnindex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("k" & 2 + columnindex).Interior.ColorIndex = 0
                    End Select
                               
                End If
                
                total = 0
                change = 0
                columnindex = columnindex + 1
                Days = 0
                dailyChange = 0
                
                Else
                
                'if the ticker is still the same add results
                    total = total + ws.Cells(rowindex, 7).Value
                    
                
            
            End If
            
            
        
        
        Next rowindex
        
        'take the max and min and place them in a separate part in the worksheet
        ws.Range("q2") = "%" & WorksheetFunction.Max(ws.Range("k2:k" & rowCount)) * 100
        ws.Range("q3") = "%" & WorksheetFunction.Min(ws.Range("k2:k" & rowCount)) * 100
        ws.Range("q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("k2:k" & rowCount)), ws.Range("k2:k" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("k2:k" & rowCount)), ws.Range("k2:k" & rowCount), 0)
        volumn_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
            
    Next ws
    
    
End Sub
