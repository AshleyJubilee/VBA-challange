Attribute VB_Name = "Module1"
Sub ticker()

    'Loops through each sheet
    Dim i As Long
    Dim shtCount As Long
    
    shtCount = Sheets.Count
    
    For i = 1 To shtCount

        Dim tickerCounter As Integer
        Dim totalVolume As LongLong
        Dim stockOpen As Double
        Dim stockClose As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As LongLong
        Dim incTicker As String
        Dim decTicker As String
        Dim volTicker As String
        Dim yearStart As Long
        Dim yearEnd As Long
    
        Dim LR As Long
        LR = Sheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    
        Sheets(i).Cells(1, 9).Value = "Ticker"
        Sheets(i).Cells(1, 10).Value = "Yearly Change"
        Sheets(i).Cells(1, 11).Value = "Percent Change"
        Sheets(i).Cells(1, 12).Value = "Total Stock Volume"
        Sheets(i).Cells(1, 15).Value = "Ticker"
        Sheets(i).Cells(1, 16).Value = "Value"
        Sheets(i).Cells(2, 14).Value = "Greatest %  Increase"
        Sheets(i).Cells(3, 14).Value = "Greatest % Decrease"
        Sheets(i).Cells(4, 14).Value = "Greatest Total Volume"
    
        totalVolume = 0
        tickerCounter = 2
        
        yearStart = 20170102 + (i * 10000)
        yearEnd = 20171231 + (i * 10000)
    
        For j = 2 To LR
        
            'If current stock matches tickerCounter, add volume
            If Sheets(i).Cells(j, 1).Value = Sheets(i).Cells(tickerCounter, 9).Value Then
            
                totalVolume = totalVolume + Sheets(i).Cells(j, 7).Value
            
            Else
        
            'If new stock, update tickerCounter and add volume
            Sheets(i).Cells(tickerCounter, 9) = Sheets(i).Cells(j, 1).Value
        
            totalVolume = totalVolume + Sheets(i).Cells(j, 7).Value
            End If
        
            'If first day of the year, set the stockOpen price
            If Sheets(i).Cells(j, 2).Value = yearStart Then
        
            stockOpen = Sheets(i).Cells(j, 3).Value
            End If
        
            'If end of year, set stockClose price, then calculate changes, and print
            If Sheets(i).Cells(j, 2).Value = yearEnd Then
        
            stockClose = Sheets(i).Cells(j, 6).Value
        
            yearlyChange = stockClose - stockOpen
        
            percentChange = yearlyChange / stockOpen
        
            Sheets(i).Cells(tickerCounter, 10) = yearlyChange
        
            Sheets(i).Cells(tickerCounter, 11) = FormatPercent(CStr(percentChange))
        
            Sheets(i).Cells(tickerCounter, 12) = totalVolume
        
                'Compares current change to greatest variables, and updates them
                If percentChange > greatestIncrease Then
            
                greatestIncrease = percentChange
                incTicker = Sheets(i).Cells(j, 1).Value
                End If
            
                If percentChange < greatestDecrease Then
            
                greatestDecrease = percentChange
                decTicker = Sheets(i).Cells(j, 1).Value
                End If
            
                If totalVolume > greatestVolume Then
            
                greatestVolume = totalVolume
                volTicker = Sheets(i).Cells(j, 1).Value
                End If
        
            'Reset volume and move to next ticker
            totalVolume = 0
            tickerCounter = tickerCounter + 1
            End If
        
        Next
        
        'Sets Conditional Formatting
        For k = 2 To LR
        
            If Sheets(i).Cells(k, 10).Value < 0 Then
            Sheets(i).Cells(k, 10).Interior.ColorIndex = 3
            Sheets(i).Cells(k, 11).Interior.ColorIndex = 3
        
            ElseIf Sheets(i).Cells(k, 10).Value > 0 Then
            Sheets(i).Cells(k, 10).Interior.ColorIndex = 4
            Sheets(i).Cells(k, 11).Interior.ColorIndex = 4
            End If
            
        Next

        Sheets(i).Cells(2, 15).Value = incTicker
        Sheets(i).Cells(2, 16).Value = CStr(FormatPercent(greatestIncrease))
        Sheets(i).Cells(3, 15).Value = decTicker
        Sheets(i).Cells(3, 16).Value = CStr(FormatPercent(greatestDecrease))
        Sheets(i).Cells(4, 15).Value = volTicker
        Sheets(i).Cells(4, 16).Value = greatestVolume
        
        Sheets(i).Columns("A:P").AutoFit
    
    Next i

End Sub



