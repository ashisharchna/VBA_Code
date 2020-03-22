Sub Stock()

    Dim lastRow As Long
    Dim curRow As Long
    Dim firstRow As Long

    Dim Total As Double
    Dim Change As Double
    
    Dim Percent_Change As Double
    Dim Greatest_Percent As Long
    Dim Lowest_Percent As Long
    
    Dim ws As Worksheet
    
    'Looping through all the worksheets
    
    For Each ws In Worksheets
    
      'Setting Initial values      
        Greatest_Percent = 0
        Lowest_Percent = 0
        curRow = 2   
        firstRow = 2
        Total = 0

        lastRow = ws.Cells(Rows.count, "A").End(xlUp).Row
            
       'looping through each cell and calculate desired values         
        For i = 2 To lastRow
            
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Setting the ticker

                    ws.Cells(curRow, 9).Value = ws.Cells(i, 1).Value

                 'Calculating Total Stock Volume   
                    Total = Total + ws.Cells(i, 7).Value
                'Assigning the total stock volume  
                    ws.Cells(curRow, 12).Value = Total
                    
                'Calculate yearly change                    
                    Change = ws.Cells(i, 6).Value - ws.Cells(firstRow, 3).Value           
                    ws.Cells(curRow, 10).Value = Change

                'conditional formatting for positive and negative change
                    If Change > 0 Then
                    
                        ws.Cells(curRow, 10).Interior.ColorIndex = 10 'green
                  
                    Else
                        ws.Cells(curRow, 10).Interior.ColorIndex = 3 'green
                    
                    End If

                'Calculate yearly percent change

                If ws.Cells(firstRow, 3).Value <> 0 Then
                    
                        Percent_Change = Change / ws.Cells(firstRow, 3).Value
                                                  
                End If
                    
               ws.Cells(curRow, 11).Value = Percent_Change
                    
                    
                    curRow = curRow + 1
                    
                    Total = 0
                    
                    firstRow = i + 1
                   
                                       
                Else
                    
                    Total = Total + ws.Cells(i, 12).Value
                    
                End If
            Next i
            
        Greatest_Percent = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        Lowest_Percent = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
            
        ws.Cells(2, 15).Value = Greatest_Percent
        ws.Cells(3, 15).Value = Lowest_Percent
             
    Next ws
        
End Sub


















