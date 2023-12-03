# stock-analysis
Sub stockChange():
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        Dim ticker_row As Long
        Dim last_row As LongLong
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim stock_volume As LongLong
        Dim open_price As Double
        
        'defined 6 variables for both ticker and value
        
        last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        open_price = ws.Cells(2, 3).Value
        stock_volume = ws.Cells(2, 7).Value
        ticker_row = 2
        
        'display column labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent  Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        For i = 2 To last_row
                If ws.Cells(i + 1, 1) = ws.Cells(i, 1) Then
                    
                    'add total volume
                    stock_volume = stock_volume + ws.Cells(i + 1, 7).Value
                
                Else
                
                    yearly_change = ws.Cells(i, 6).Value - open_price
                    percent_change = yearly_change / open_price
                       
                    'display results to summary table
                    ws.Range("I" & ticker_row).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & ticker_row).Value = yearly_change
                    ws.Range("K" & ticker_row).Value = percent_change
                    ws.Range("L" & ticker_row).Value = stock_volume
                    
                    'reset values
                    stock_volume = ws.Cells(i + 1, 7).Value
                    open_price = ws.Cells(i + 1, 3).Value
                    ticker_row = ticker_row + 1
                    
               End If
        
        Next i
       
       'new for loop to caculate gpinc, gpdec, and gtvol
           For K = 2 To ticker_row
            
                Dim gp_increase As Double
                Dim gp_decrease As Double
                Dim gt_volume As LongLong
                
               gp_increase = Application.WorksheetFunction.max(Range("K:K"))
               gp_decrease = Application.WorksheetFunction.min(Range("K:K"))
               gt_volume = Application.WorksheetFunction.max(Range("L:L"))
               
               If Cells(K, 11).Value = gp_increase Then
                    Cells(2, 16).Value = gp_increase
                    Cells(2, 15).Value = Cells(K, 9).Value
                
                ElseIf Cells(K, 11).Value = gp_decrease Then
                    Cells(3, 16).Value = gp_decrease
                    Cells(3, 15).Value = Cells(K, 9).Value
                
               ElseIf Cells(K, 12).Value = gt_volume Then
                    Cells(4, 16).Value = gt_volume
                    Cells(4, 15).Value = Cells(K, 9).Value
               End If
            
           Next K
           
           For j = 2 To last_row2
           
                Dim last_rowJ As Integer
                
                last_rowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
                    
                If Cells(j, 10).Value >= 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                Else:
                    Cells(j, 10).Interior.ColorIndex = 3
                
                End If
            
            Next j
            
            For s = 2 To last_row3
            
                Dim last_rowP As Integer
                
                last_rowP = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
                
                Cells(s, 11).Value = FormatPercent(Cells(s, 11), 2, vbFalse, vbFalse, vbFalse)
            
            Next s
            
            For x = 2 To 3
                
                Cells(x, 16).Value = FormatPercent(Cells(x, 16), 2, vbFalse, vbFalse, vbFalse)
            
            Next x
                
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Value"
            Cells(2, 14).Value = "Greatest % Increase"
            Cells(3, 14).Value = "Greatest % Decrease"
            Cells(4, 14).Value = "Greatest Total Volume"
        
    Next ws

End Sub
