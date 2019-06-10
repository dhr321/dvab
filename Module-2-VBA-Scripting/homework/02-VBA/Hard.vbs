Sub AnnualVolume()
    Dim i, j As LongLong
    Dim currVol As LongLong
    Dim greatestVol As LongLong
    
    'Initialize the variables.
    i = 2
    j = 2
    currTicker = Cells(i, 1).Value
    currVol = 0
    greatestVol = 0
    yearopen = Cells(i, 3).Value
    pctInc = 0
    greatestPctInc = 0
    greatestPctDec = 0
    
    
    'Go through every row until blank in first cell. Number of rows unknown.
    Do While Cells(i, 1).Value <> ""
    
        'Sum the volume for each ticker symbol until the ticker symbol changes.
        Do While currTicker = Cells(i, 1).Value
            currVol = currVol + Cells(i, 7).Value
            yearclose = Cells(i, 6).Value
            i = i + 1
            
        Loop
        
        'Populate the cells with the year-end totals.
        Cells(j, 9).Value = currTicker
        
        'Check for division by zero. Some of tickers have zeroes for all values in the data set.
        If yearclose <> 0 Then
            Cells(j, 10).Value = yearclose - yearopen
            Cells(j, 11).Value = (yearclose - yearopen) / yearopen
            Cells(j, 12).Value = currVol
        Else
            Cells(j, 10).Value = 0
            Cells(j, 11).Value = 0
            Cells(j, 12).Value = 0
        End If
         
        'Color the negative changes red and the positive changes green.
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'Find the greatest increase and decrease in percent change.
        pctInc = Cells(j, 11).Value
        
        If pctInc > greatestPctInc Then
            greatestPctInc = pctInc
            greatestPctIncTicker = currTicker
        ElseIf pctInc < greatestPctDec Then
            greatestPctDec = pctInc
            greatestPctDecTicker = currTicker
        End If
        
        'Find the greatest total stock volume.
        If currVol > greatestVol Then
            greatestVol = currVol
            greatestTicker = currTicker
        End If
        
        
        'Increment the rows in the year-end totals.
        j = j + 1
        
        'Set a new value for the current ticker symbol.
        currTicker = Cells(i, 1).Value
        
        'Reset year open.
        yearopen = Cells(i, 3).Value
                
        'Reset the volume for the next ticker symbol.
        currVol = 0
        
    Loop
    
    'Populate the cells with the Greatest values
    Cells(2, 16).Value = greatestPctIncTicker
    Cells(2, 17).Value = greatestPctInc
    Cells(3, 16).Value = greatestPctDecTicker
    Cells(3, 17).Value = greatestPctDec
    Cells(4, 16).Value = greatestTicker
    Cells(4, 17).Value = greatestVol
    
    
End Sub







