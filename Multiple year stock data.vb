Sub Alpha():

    For Each WS In Worksheets
    
    
        Dim stockName As String
        
        Dim stockVolume As Double
        stockVolume = 0
        
        Dim stockOpen As Double
        
        Dim stockClose As Double

        Dim stockValueChange As Double
        
        Dim stockPercentChange As Double
        
        Dim tickerCount As Double
        tickerCount = 0
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = WS.Cells(1, Columns.Count).End(xlToLeft).Column
    
        Dim Summary_Table As Integer
        Summary_Table = 2
        WS.Cells(1, 9).Value = "Stock Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Volume"

        Dim Summary_Table2 As Integer
        Summary_Table2 = 3
        WS.Cells(1, 15).Value = ""
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"


        For i = 2 To LastRow
            
                If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                
                    stockName = WS.Cells(i, 1).Value
                    stockVolume = stockVolume + WS.Cells(i, 7).Value
                    stockOpen = WS.Cells(i - tickerCount, 3).Value
                    stockClose = WS.Cells(i, 6).Value
                    stockValueChange = stockClose - stockOpen
                    
                        If stockOpen = 0 Then
                            stockPercentageChange = stockValueChange / Null
                        Else
                            stockPercentChange = stockValueChange / stockOpen
                    
                        End If
                                    
                    WS.Range("I" & Summary_Table).Value = stockName
                    WS.Range("J" & Summary_Table).Value = stockValueChange
                    
                        If WS.Range("J" & Summary_Table).Value < 0 Then
                            WS.Range("J" & Summary_Table).Interior.Color = vbRed
                        Else
                            WS.Range("J" & Summary_Table).Interior.Color = vbGreen
                        End If
                    
                    WS.Range("K" & Summary_Table).Value = stockPercentChange
                    WS.Range("K" & Summary_Table).NumberFormat = "0.00%"
                    WS.Range("L" & Summary_Table).Value = stockVolume
                    
                    
                    
                    Summary_Table = Summary_Table + 1
                    stockVolume = 0
                    tickerCount = 0
                    
                Else
                    
                    stockVolume = Stock_Volume + WS.Cells(i, 7).Value
                    tickerCount = tickerCount + 1
                    
                End If
                
                
        Next i
        
    Next WS
    
End Sub
