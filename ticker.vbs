Sub ticker()
    Dim lastRow As Long
    Dim curRow As Long
    Dim Total As Double
    Dim Change As Double
    Dim firstRow As Long
    
    curRow = 2
    
    firstRow = 2
    
    Total = 0
    lastRow = Cells(Rows.count, "A").End(xlUp).Row
    
   'MsgBox ("Total rows are " & lastRow)
   
    For i = 2 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Cells(curRow, 9).Value = Cells(i, 1).Value
            
            Total = Total + Cells(i, 7).Value
            
            Cells(curRow, 10).Value = Total
            
            'MsgBox ("current Row is" & curRow & " and Total volume is" & Total)
            
            Change = Cells(i, 6).Value - Cells(firstRow, 3).Value
            
            Cells(curRow, 11).Value = Change
            
            
            curRow = curRow + 1
            
            Total = 0
            
            firstRow = i + 1
            
            
            
        Else
            
            Total = Total + Cells(i, 7).Value
            
        End If
    Next i
End Sub