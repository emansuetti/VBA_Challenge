Attribute VB_Name = "Module1"
'=======================
'The VBA of Wall Street!
'=======================

Sub VBAofWallStreet()


'Delcaring all of our variables for the project

        Dim Ticker As String

        Dim SummaryTable As Double
    
                SummaryTable = 2
    
        Dim TotalVolume As Double
        
'Note* tried to declare TotalVolume as long but kept getting error message. I guess long was too short...

                TotalVolume = 0
    
        Dim LastRow As Double

                LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim YearlyChange As Double
    
        Dim StockOpen As Double
    
                StockOpen = Cells(2, 3).Value
    
        Dim StockClose As Double
                
      
        Dim PercentageChange As Double
    
    
'Time to set our columns needed for our new sorted and calculated information

    Cells(1, 10).Value = "Ticker Name"
    
    Cells(1, 11).Value = "Yearly Change"
    
    Cells(1, 12).Value = "Percentage Change"
    
    Cells(1, 13).Value = "Total Volume"
        
'Starting the loop
    
For i = 2 To LastRow

'Looking for same ticker name in Column "A" if not add values and move to next Ticker and add name to Table
'with TotalVolume calculation

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
        
        Range("J" & SummaryTable).Value = Ticker
    
        Range("M" & SummaryTable).Value = TotalVolume
                                
' Create value for StockClose inside loop
            
            StockClose = Cells(i, 6).Value
            
' Create Value for YearlyChange and place value in Table

            YearlyChange = (StockClose - StockOpen)
            
        Range("K" & SummaryTable).Value = YearlyChange
        
        
'Creating color codes for positive and negative outcomes (tried to do this at the end of the script but would not work as intended)
        

    If Range("K" & SummaryTable).Value <= 0 Then
    
        Range("K" & SummaryTable).Interior.ColorIndex = 3
         
            Else
            
        Range("K" & SummaryTable).Interior.ColorIndex = 4

     End If

'In order to not get an error for this next line of code I had to add the And function that stopped the
'error finally after some debugging and a little bit of luck

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And StockOpen <> 0 Then
    
            PercentageChange = YearlyChange / StockOpen
          
    End If
    
'Format for Percentages and Decimal Places in Table and reseting values for next loop

        Range("L" & SummaryTable).NumberFormat = "0.00%"
        
        Range("L" & SummaryTable).Value = PercentageChange
                
'
                SummaryTable = SummaryTable + 1
            
                TotalVolume = 0
            
                StockOpen = Cells(i + 1, 3).Value
    
                StockClose = Cells(i, 6).Value
    Else
            
    
            TotalVolume = TotalVolume + Cells(i, 7).Value
      
    End If
    

    
Next i

                    
                    
End Sub
