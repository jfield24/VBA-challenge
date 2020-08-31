Sub multipleyearstockdata()

'Sets the worksheet as a variable'

    Dim ws As Worksheet
    
'Need to use a forloop through every worksheet'
        For Each ws In Worksheets
    
    'Sets variable for holding the ticker name'
            Dim ticker As String
            
    'Sets initial variable for holding the opening price'
    
            Dim openingprice As Double
            openingprice = 0
            
    'Sets initial variable for holding the closing price'
    
            Dim closingprice As Double
            closingprice = 0
            
    'Sets initial variable for holding the yearly change'
            Dim yearlychange As Double
            yearlychange = 0
        
    'Sets initial variable for holding the percentage change'
            Dim percentchange As Double
            percentchange = 0
            
    'Sets initial variable for holding the total volume'
            Dim volume As Double
            volume = 0
           
    'Sets initial variable for holding the value of the stock with greatest increase'
            Dim maxincrease As Double
            maxincrease = 0
           
    'Sets variable for holding the ticker name of stock with greatest increase'
            Dim increaseticker As String
            
    'Sets initial variable for holding the value of the stock with greatest decrease'
            Dim maxdecrease As Double
            maxdecrease = 0
            
    'Sets variable for holding the ticker name of stock with greatest decrease'
            Dim decreaseticker As String
            
    'Sets initial variable for holding the value of the stock with greatest volume'
            Dim maxvolume As Double
            maxvolume = 0
            
     'Sets variable for holding the ticker name of stock with greatest volume'
            Dim maxvolticker As String
     
     'Keeps track of location of the outputs in the summary table'
     
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
      'Determines the last row'
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
       'Sets the headers of the summary table'
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
       'Tells the forloop where the first opening price value in the data is'
            openingprice = ws.Cells(2, 3).Value
        
        'Establishes where we need to loop'
            For i = 2 To lastrow
            
                
               'If statement to tell us to only add to summary table when we change ticker name'
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                
                'Locates ticker name and puts it in the summary table'
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & Summary_Table_Row).Value = ticker
            
                'Locates closing price of the stock'
                    closingprice = ws.Cells(i, 6).Value
                    
                'Computes the yearly change for the stock and puts it in the summary table'
                    yearlychange = closingprice - openingprice
                    ws.Range("J" & Summary_Table_Row).Value = yearlychange
                        
                    'If statement says if yearly change is positive then colours the cell green, if negative red'
                        If (yearlychange > 0) Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf (yearlychange < 0) Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                      
                    'Kept getting divide by zero error - so made an if statement to make it only calculate percent change if opening price isn't 0'
                        If openingprice <> 0 Then
                            percentchange = (yearlychange / openingprice)
                        
                        End If
                    
                'Puts the computed percentage change of the stock in the summary table, and converts to a percentage'
                    ws.Range("K" & Summary_Table_Row).Value = percentchange
                    ws.Range("K" & Summary_Table_Row).Style = "Percent"
                    
                    
                 'Computes total volume for the stock, and adds the total for that stock in the summary table'
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Range("L" & Summary_Table_Row).Value = volume
                    
                
                'Tells the forloop to add 1 to the summary table row count'
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                      
                    'If statement to find the stock with greatest increase and the relevant stock ticker'
                        If (percentchange > maxincrease) Then
                            maxincrease = percentchange
                            increaseticker = ticker
                            
                     'Else If statement to find the stock with greatest decrease and the relevant stock ticker'
                        ElseIf (percentchange < maxdecrease) Then
                            maxdecrease = percentchange
                            decreaseticker = ticker
                        
                        End If
                        
                'Puts the greatest % increase and the stock ticker it is for in the summary table, converts the value of the increase to %'
                    ws.Range("P2").Value = increaseticker
                    ws.Range("Q2").Value = maxincrease
                    ws.Range("Q2").Style = "Percent"
                    
                'Puts the greatest % decrease and the stock ticker it is for in the summary table, converts the value of the decrease to %'
                    ws.Range("P3").Value = decreaseticker
                    ws.Range("Q3").Value = maxdecrease
                    ws.Range("Q3").Style = "Percent"
                    
                    'If statement to find the stock with greatest total volume and the relevant stock ticker'
                        If (volume > maxvolume) Then
                            maxvolume = volume
                            maxvolticker = ticker
                            
                        End If
                    
                'Puts the greatest total volume and the stock ticker it is for in the summary table'
                    ws.Range("P4").Value = maxvolticker
                    ws.Range("Q4").Value = maxvolume
                    
                   
                'Resets the value of the variables'
                    openingprice = ws.Cells(i + 1, 3).Value
                    closingprice = 0
                    yearlychange = 0
                    percentchange = 0
                    volume = 0
                    
                'Else statement says that if stock ticker is same as previous row, add to the total volume for that stock'
                Else: volume = volume + ws.Cells(i, 7).Value
                
                End If
            
            'Allows the forloop to work by looping to the next i value'
            Next i
        
      'Allows the forloop to work by looping to the next worksheet'
        Next ws
            

End Sub