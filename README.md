'Code for Stock data analysis for module 2 challenge

Sub stock_analysis()

'Define variables
Dim row_index As Long
Dim col_index As Integer
Dim Total_stock_volume As Variant
Dim Yearly_change As Variant
Dim Percent_change As Variant
Dim start As Long
Dim counter_row As Long
Dim days As Integer
Dim Daily_change As Single
Dim ws As Worksheet


For Each ws In Worksheets

        Total_stock_volume = 0
        col_index = 0
        start = 2
        Yearly_change = 0
        find_value = 0
        

        'setting labels for each cell

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly_change"
        ws.Range("K1").Value = "Percent_change"
        ws.Range("L1").Value = "Total_stock_volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

'find out the number of rows in each worksheet
counter_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For row_index = 2 To counter_row

        If ws.Cells(row_index + 1, 1).Value <> ws.Cells(row_index, 1).Value Then

                Total_stock_volume = Total_stock_volume + ws.Cells(row_index, 7).Value
    
                             If Total_stock_volume = 0 Then
                                          ws.Range("I" & 2 + col_index).Value = Cells(row_index, 1).Value
                                          ws.Range("J" & 2 + col_index).Value = 0
                                          ws.Range("K" & 2 + col_index).Value = "%" & 0
                                          ws.Range("L" & 2 + col_index).Value = 0
    
                             Else
        
                                     If ws.Cells(start, 3) = 0 Then
                                           For find_value = start To row_index
                                                  If ws.Cells(find_value, 3).Value <> 0 Then
                                                        start = find_value
                                                        Exit For
                                                 End If
                                            Next find_value
                                       End If
       
                                 Yearly_change = (ws.Cells(row_index, 6) - ws.Cells(start, 3))
                                 Percent_change = Yearly_change / ws.Cells(start, 3)
        
                                start = row_index + 1
      
                                ws.Range("I" & 2 + col_index) = ws.Cells(row_index, 1).Value
                                ws.Range("J" & 2 + col_index) = Yearly_change
                                ws.Range("J" & 2 + col_index).NumberFormat = "0.00"
                                ws.Range("K" & 2 + col_index).Value = Percent_change
                                ws.Range("K" & 2 + col_index).NumberFormat = "0.00%"
                                ws.Range("L" & 2 + col_index).Value = Total_stock_volume

                                 Select Case Percent_change
                                        Case Is > 0
                                             ws.Range("J" & 2 + col_index).Interior.ColorIndex = 4
                                        Case Is < 0
                                              ws.Range("J" & 2 + col_index).Interior.ColorIndex = 3
                                        Case Else
                                             ws.Range("J" & 2 + col_index).Interior.ColorIndex = 0
                                 End Select
         
                            End If

                        Total_stock_volume = 0
                        Yearly_change = 0
                        col_index = col_index + 1
         
    
         Else
                    Total_stock_volume = Total_stock_volume + ws.Cells(row_index, 7).Value
        
        End If
    
    Next row_index
    
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & counter_row)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & counter_row)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & counter_row))
    
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & counter_row)), ws.Range("K2:K" & counter_row), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & counter_row)), ws.Range("K2:K" & counter_row), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & counter_row)), ws.Range("L2:L" & counter_row), 0)


   ws.Range("P2") = ws.Cells(increase_number + 1, 9)
   ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
   ws.Range("P4") = ws.Cells(volume_number + 1, 9)


Next ws


End Sub
