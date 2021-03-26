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
'Note* tried to declare TotalVolume as long but kept getting error message
                TotalVolume = 0
    
        Dim LastRow As Double

                LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim YearlyChange As Double
    
        Dim StockOpen As Double
    
                StockOpen = Cells(2, 3).Value
    
        Dim StockClose As Double
      
                StockClose = Cells(2, 6).Value
                
        Dim PercentageChange As Double
    
    
'Time to set our columns needed for our new sorted and calculated information

    Cells(1, 10).Value = "Ticker Name"
    
    Cells(1, 11).Value = "Yearly Change"
    
    Cells(1, 12).Value = "Percentage Change"
    
    Cells(1, 13).Value = "Total Volume"
        
'Starting the loop
    
For i = 2 To LastRow

'Looking for same ticker name in Column "A" if not add values and move to next Ticker and add name to Table

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
        
        Range("J" & SummaryTable).Value = Ticker
    
        Range("M" & SummaryTable).Value = TotalVolume
                                
' Create value for StockClose
            StockOpen = Cells(i, 3).Value
            
            
            StockClose = Cells(i, 6).Value
            
' Create Value for YearlyChange and place value in Table

            YearlyChange = (Cells(i, 6).Value - StockOpen)
            
        Range("K" & SummaryTable).Value = YearlyChange

' Broken peice of code....D: cant create secondare else without error and does not generate PercentageChange as intended

    
    PercentageChange = YearlyChange / StockOpen
        
    
'Format for Percentages and Decimal Places in Table

        Range("L" & SummaryTable).NumberFormat = "0.00%"
        
        Range("L" & SummaryTable).Value = PercentageChange
                
'
                SummaryTable = SummaryTable + 1
            
                TotalVolume = 0
            
                StockOpen = Cells(i + 1, 3).Value
    
    Else
    
            StockClose = Cells(i, 6).Value
                        
            Cells(i, 15).Value = StockClose
            
            Cells(i, 14).Value = StockOpen
    
            TotalVolume = TotalVolume + Cells(i, 7).Value
      
    End If
    
Next i

   
                    
End Sub
