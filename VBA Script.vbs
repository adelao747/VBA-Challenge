Attribute VB_Name = "Module1"
Sub stocks_report()
    
    ' Names of rows and headers
    
        For Each ws In Worksheets
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
        
    ' Ticker symbols, yearly change at beginning to close, percentage change from opening to closing and total stock volume.
        Dim ticker_symbol As String
        Dim ticker_volume, opening_price, closing_price, yearly_change, percent_change As Double
        Dim ticker_row As Integer
        
        ticker_volume = 0
        ticker_row = 2
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker_symbol = ws.Cells(i, 1).Value
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            ws.Range("I" & ticker_row).Value = ticker_symbol
            ws.Range("L" & ticker_row).Value = ticker_volume
            closing_price = ws.Cells(i, 6).Value
            yearly_change = (closing_price - opening_price)
            ws.Range("J" & ticker_row).Value = yearly_change
            
                If opening_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / opening_price
                End If
                
            ws.Range("K" & ticker_row).Value = percent_change
            ws.Range("K" & ticker_row).Style = "Percent"
            
            ticker_row = ticker_row + 1
            ticker_volume = 0
            opening_price = ws.Cells(i + 1, 3)
            
            Else
            
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                
            End If
            
            Next i
            
        lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Conditional formatting
            For i = 2 To lastrow_summary_table
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 10
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i
            
            
        ' Greatest % increase, greatest % decrease and greatest total volume
            Dim increase_ticker, decrease_ticker, total_ticker As String
            
            Greatest_Increase = ws.Range("K2").Value
            Greatest_Decrease = ws.Range("K2").Value
            Greatest_Total = ws.Range("L2").Value
         

         
         Lastrow_Ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row
         For r = 2 To Lastrow_Ticker:
               If ws.Range("K" & r + 1).Value > Greatest_Increase Then
                  Greatest_Increase = ws.Range("K" & r + 1).Value
                  Greatest_Increase_Ticker = ws.Range("I" & r + 1).Value
               ElseIf ws.Range("K" & r + 1).Value < Greatest_Decrease Then
                  Greatest_Decrease = ws.Range("K" & r + 1).Value
                  Greatest_Decrease_Ticker = ws.Range("I" & r + 1).Value
                ElseIf ws.Range("L" & r + 1).Value > Greatest_Total Then
                  Greatest_Total = ws.Range("L" & r + 1).Value
                  Greatest_Total_Ticker = ws.Range("I" & r + 1).Value
                End If
            Next r
            
         
                    ws.Range("P2").Value = Greatest_Increase_Ticker
                    ws.Range("P3").Value = Greatest_Decrease_Ticker
                    ws.Range("P4").Value = Greatest_Total_Ticker
                    ws.Range("Q2").Value = Greatest_Increase
                    ws.Range("Q3").Value = Greatest_Decrease
                    ws.Range("Q4").Value = Greatest_Total
                    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
            
    
 Next ws
 
            
        
End Sub


