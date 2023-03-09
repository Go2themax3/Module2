Sub StockAnalysis()
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Double
    Dim stockVolume As Double
    Dim i As Long
    Dim j As Long
    stockVolume = 0
    
    j = 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
       
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        openingPrice = Cells(i, 3).Value
        closingPrice = Cells(i, 6).Value
        stockVolume = Cells(i, 7).Value + stockVolume
        yearlyChange = closingPrice - openingPrice
        percentChange = yearlyChange / openingPrice
       

       
    
        Range("H" & j + 2).Value = Cells(i, 1).Value
        Range("I" & j + 2).Value = yearlyChange
        Range("J" & j + 2).Value = percentChange
        Range("K" & j + 2).Value = stockVolume
        j = j + 1
        
      
        stockVolume = 0
        Else
        stockVolume = Cells(i, 7).Value + stockVolume
      
     End If
     
    Next i
    
    
    
End Sub

