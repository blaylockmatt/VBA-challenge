Sub stockanalysis():

    'Declarations
    Dim total As Double
    Dim RowCount As Long
    Dim summary_row As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim start As Long
    Dim Greatest_Increase As Double
    Dim Lowest_Increase As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
       
    
    
    start = 2
    total = 0
    'find the number of total rows
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Lowest % Increase"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(5, 15).Value = "Least Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    summary_row = 2
    
        For i = 2 To RowCount
          If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                total = total + ws.Cells(i, 7).Value
               
                
                If ws.Cells(start, 3) = 0 Then
                    For Value = start To i
                        If ws.Cells(Value, 3).Value <> 0 Then
                            start = Value
                            Exit For
                        End If
                    Next Value
                End If
                yearly_change = ws.Cells(i, 6) - ws.Cells(start, 3)
                percent_change = Round((yearly_change / ws.Cells(start, 3)) * 100, 2)
                ws.Range("I" & summary_row).Value = ws.Cells(i, 1).Value
                ws.Range("L" & summary_row).Value = total
                ws.Range("J" & summary_row).Value = yearly_change
                ws.Range("K" & summary_row).Value = percent_change
                If yearly_change > 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                End If
                    
                               
                start = i + 1
                        
                total = 0
                summary_row = summary_row + 1
            Else
                total = total + ws.Cells(i, 7).Value
                
            End If
            
                   
            
    Next i
    Next ws
End Sub
