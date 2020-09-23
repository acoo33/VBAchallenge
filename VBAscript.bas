Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet
For Each ws In Worksheets

'new column headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest Percent Increase"
    ws.Cells(3, 16).Value = "Greatest Percent Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
 
 'format cosider center format
    ws.Range("J1:M1").Font.Bold = True
    ws.Range("Q1:R1").Font.Bold = True
    ws.Range("P2:P4").Font.Bold = True
    ws.Range("J1:M1").HorizontalAlignment = xlCenter
    ws.Range("Q1:R1").HorizontalAlignment = xlCenter
    ws.Range("L:L").NumberFormat = "0.00%"
    ws.Range("R2:R3").NumberFormat = "0.00%"
    
  '# rows and columns
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Initiate Variables
    opening = ws.Range("C2").Value
    counter = 2
    total_v = 0
    great_incr = 0
    great_decr = 0
    total_v_control = 0
    
    For i = 2 To lastrow
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        total_v = total_v + volume
        
        
            If ticker <> ws.Cells(i + 1, 1).Value Then
                closing = ws.Cells(i, 6).Value
                ws.Cells(counter, 10).Value = ticker
                
                y_change = closing - opening
                ws.Cells(counter, 11).Value = y_change
                
                If y_change > 0 Then
                    ws.Cells(counter, 11).Interior.Color = vbGreen
                ElseIf y_change < 0 Then
                    ws.Cells(counter, 11).Interior.Color = vbRed
                End If
                
            
                    If opening = 0 Then
                        p_change = 0
                    Else
                        p_change = y_change / opening
                        ws.Cells(counter, 12).Value = p_change
                        
                        If great_incr < p_change Then
                            great_incr = p_change
                            great_incr_ticker = ticker
                        ElseIf great_decr > p_change Then
                            great_decr = p_change
                            great_decr_ticker = ticker
                        End If
                        
                        
                    End If
                
                ws.Cells(counter, 13).Value = total_v
                
                    If total_v > total_v_control Then
                        total_v_control = total_v
                        total_v_ticker = ticker
                    End If
                
                
                ' Updating variables
                opening = ws.Cells(i + 1, 3).Value
                counter = counter + 1
                total_v = 0
             End If
       
    Next i
        
    ws.Range("Q2").Value = great_incr_ticker
    ws.Range("Q3").Value = great_decr_ticker
    ws.Range("R2").Value = great_incr
    ws.Range("R3").Value = great_decr
    ws.Range("Q4").Value = total_v_ticker
    ws.Range("R4").Value = total_v_control

Next ws

End Sub
